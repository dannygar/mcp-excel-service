"""
MCP Excel Service - Streamable HTTP Transport

This server provides Excel manipulation tools via the Model Context Protocol (MCP).
Uses Microsoft Graph API to interact with Excel files in SharePoint/OneDrive.
Designed for deployment on Azure Container Apps with Foundry Agent integration.

Authentication:
    Uses Azure AD service principal (client credentials flow) for Graph API access.
    Required environment variables:
    - AZURE_TENANT_ID: Azure AD tenant ID
    - AZURE_CLIENT_ID: App registration client ID
    - AZURE_CLIENT_SECRET: App registration client secret
"""

import os
import json
import logging
import time
import re
from datetime import datetime, timedelta
from typing import Optional, Union

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse

# Import configuration
from config import (
    TRADE_TRACKER_URL,
    TRADE_TRACKER_FILE,
    map_strategy_name,
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mcp-excel-server")

# Load environment variables
load_dotenv()

# Initialize MCP server
mcp = FastMCP("MCP Excel Service")

# Microsoft Graph API base URL
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"


# Token cache for client credentials flow
_token_cache = {
    "access_token": None,
    "expires_at": 0,
}


# =============================================================================
# Helper Functions
# =============================================================================

def parse_date_string(value: str) -> datetime | None:
    """
    Try to parse a string as a date using common formats.
    
    Supported formats:
    - MM/DD/YYYY (e.g., 12/22/2025)
    - YYYY-MM-DD (e.g., 2025-12-22)
    - M/D/YYYY (e.g., 1/5/2025)
    - DD-MM-YYYY (e.g., 22-12-2025)
    
    Returns:
        datetime object if parsing succeeds, None otherwise
    """
    date_formats = [
        "%m/%d/%Y",  # 12/22/2025
        "%Y-%m-%d",  # 2025-12-22
        "%d-%m-%Y",  # 22-12-2025
        "%m-%d-%Y",  # 12-22-2025
        "%Y/%m/%d",  # 2025/12/22
    ]
    
    for fmt in date_formats:
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    
    return None


def date_to_excel_serial(dt: datetime) -> int:
    """
    Convert a datetime object to an Excel serial number.
    
    Excel serial numbers count days since January 1, 1900 (with a bug that
    treats 1900 as a leap year, so dates after Feb 28, 1900 are off by 1).
    
    Args:
        dt: datetime object to convert
    
    Returns:
        Excel serial number as integer
    """
    # Excel's epoch is December 30, 1899 (to account for the leap year bug)
    excel_epoch = datetime(1899, 12, 30)
    delta = dt - excel_epoch
    return delta.days


def excel_serial_to_date(serial: Union[int, float]) -> datetime:
    """
    Convert an Excel serial number to a datetime object.
    
    Args:
        serial: Excel serial number
    
    Returns:
        datetime object
    """
    excel_epoch = datetime(1899, 12, 30)
    return excel_epoch + timedelta(days=int(serial))


def is_likely_date_string(value: str) -> bool:
    """
    Check if a string looks like a date format.
    
    Returns:
        True if the string matches common date patterns
    """
    # Pattern for common date formats: M/D/YYYY, MM/DD/YYYY, YYYY-MM-DD, etc.
    date_patterns = [
        r'^\d{1,2}/\d{1,2}/\d{4}$',  # M/D/YYYY or MM/DD/YYYY
        r'^\d{4}-\d{2}-\d{2}$',       # YYYY-MM-DD
        r'^\d{2}-\d{2}-\d{4}$',       # DD-MM-YYYY or MM-DD-YYYY
        r'^\d{4}/\d{2}/\d{2}$',       # YYYY/MM/DD
    ]
    
    for pattern in date_patterns:
        if re.match(pattern, value.strip()):
            return True
    return False


def compare_values_for_search(cell_value, reference_value: str) -> bool:
    """
    Compare a cell value with a reference value for search, handling date conversions.
    
    If the reference_value looks like a date string and the cell contains a number
    that could be an Excel serial date, convert and compare as dates.
    
    Args:
        cell_value: The value from the Excel cell (could be number, string, etc.)
        reference_value: The value to search for (always a string)
    
    Returns:
        True if values match, False otherwise
    """
    if cell_value is None:
        return False
    
    # Direct string comparison first
    if str(cell_value) == str(reference_value):
        return True
    
    # Check if reference_value looks like a date
    if is_likely_date_string(reference_value):
        parsed_date = parse_date_string(reference_value)
        if parsed_date:
            # Convert to Excel serial number
            reference_serial = date_to_excel_serial(parsed_date)
            
            # Check if cell_value is a number (potential Excel date serial)
            try:
                cell_serial = float(cell_value)
                # Excel dates are typically > 1 and < 2958465 (year 9999)
                if 1 <= cell_serial <= 2958465:
                    # Compare as integers (dates are whole numbers)
                    if int(cell_serial) == reference_serial:
                        return True
            except (ValueError, TypeError):
                pass
    
    # Check if cell_value is a number that might be an Excel date serial
    # and reference_value is also numeric
    try:
        cell_num = float(cell_value)
        ref_num = float(reference_value)
        if cell_num == ref_num:
            return True
    except (ValueError, TypeError):
        pass
    
    return False


def get_client_credentials() -> tuple[str, str, str]:
    """Get Azure AD client credentials from environment."""
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    
    if not all([tenant_id, client_id, client_secret]):
        missing = []
        if not tenant_id:
            missing.append("AZURE_TENANT_ID")
        if not client_id:
            missing.append("AZURE_CLIENT_ID")
        if not client_secret:
            missing.append("AZURE_CLIENT_SECRET")
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")
    
    return tenant_id, client_id, client_secret


async def get_access_token() -> str:
    """
    Get Microsoft Graph API access token using client credentials flow.
    
    Implements token caching to avoid unnecessary token requests.
    Tokens are refreshed 5 minutes before expiration.
    """
    global _token_cache
    
    # Check if we have a valid cached token (with 5-minute buffer)
    if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 300:
        return _token_cache["access_token"]
    
    logger.info("Acquiring new access token via client credentials flow")
    
    tenant_id, client_id, client_secret = get_client_credentials()
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.post(
            token_url,
            data=token_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30.0,
        )
        
        if response.status_code != 200:
            error_data = response.json() if response.content else {}
            error_desc = error_data.get("error_description", response.text)
            raise ValueError(f"Failed to acquire token: {error_desc}")
        
        token_response = response.json()
        
        # Cache the token
        _token_cache["access_token"] = token_response["access_token"]
        _token_cache["expires_at"] = time.time() + token_response.get("expires_in", 3600)
        
        logger.info("Successfully acquired new access token")
        return _token_cache["access_token"]


async def get_graph_headers() -> dict:
    """Get headers for Microsoft Graph API requests."""
    token = await get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def build_workbook_url(
    drive_id: str,
    item_id: str,
    site_id: Optional[str] = None,
) -> str:
    """
    Build the base URL for workbook operations.
    
    Args:
        drive_id: The ID of the drive containing the Excel file
        item_id: The ID of the Excel file item
        site_id: Optional SharePoint site ID (for SharePoint-hosted files)
    
    Returns:
        Base URL for workbook operations
    """
    if site_id:
        return f"{GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}/workbook"
    return f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}/workbook"


def parse_sharepoint_url(url: str) -> dict:
    """
    Parse a SharePoint or OneDrive URL to extract components.
    
    Supports formats:
    - https://{tenant}.sharepoint.com/sites/{sitename}/Shared%20Documents/{path}
    - https://{tenant}.sharepoint.com/Shared%20Documents/{path}
    - https://{tenant}-my.sharepoint.com/personal/{user}/Documents/{path}
    
    Returns:
        Dictionary with hostname, site_path, and file_path
    """
    from urllib.parse import urlparse, unquote
    
    parsed = urlparse(url)
    hostname = parsed.hostname
    path = unquote(parsed.path)
    
    # Remove trailing slashes and query params
    path = path.rstrip('/')
    
    # Common document library names and their variations
    doc_lib_patterns = [
        '/Shared Documents/',
        '/Documents/',
        '/Shared%20Documents/',
    ]
    
    site_path = ""
    file_path = ""
    
    # Check if this is a /sites/ or /teams/ URL
    if '/sites/' in path or '/teams/' in path:
        # Extract site path (e.g., /sites/MySite)
        parts = path.split('/')
        for i, part in enumerate(parts):
            if part in ('sites', 'teams') and i + 1 < len(parts):
                site_path = f"/{part}/{parts[i + 1]}"
                # Find the document library and file path
                remaining = '/'.join(parts[i + 2:])
                for pattern in doc_lib_patterns:
                    clean_pattern = pattern.strip('/')
                    if remaining.startswith(clean_pattern):
                        file_path = remaining[len(clean_pattern):].lstrip('/')
                        break
                    elif '/' in remaining:
                        # The first segment after site is the library
                        lib_and_path = remaining.split('/', 1)
                        if len(lib_and_path) > 1:
                            file_path = lib_and_path[1]
                break
    elif '/personal/' in path:
        # OneDrive for Business: /personal/{user}/Documents/{path}
        parts = path.split('/')
        for i, part in enumerate(parts):
            if part == 'personal' and i + 1 < len(parts):
                site_path = f"/personal/{parts[i + 1]}"
                remaining = '/'.join(parts[i + 2:])
                for pattern in doc_lib_patterns:
                    clean_pattern = pattern.strip('/')
                    if remaining.startswith(clean_pattern):
                        file_path = remaining[len(clean_pattern):].lstrip('/')
                        break
                break
    else:
        # Root site with document library
        for pattern in doc_lib_patterns:
            clean_pattern = pattern.strip('/')
            if clean_pattern in path:
                idx = path.find(clean_pattern)
                file_path = path[idx + len(clean_pattern):].lstrip('/')
                break
    
    # Handle Forms/AllItems.aspx view URLs - extract the folder path
    if 'Forms/AllItems.aspx' in file_path:
        file_path = file_path.replace('Forms/AllItems.aspx', '').rstrip('/')
    
    return {
        "hostname": hostname,
        "site_path": site_path,
        "file_path": file_path,
    }


async def resolve_excel_file_ids(
    url: str,
    file_name: str,
) -> dict:
    """
    Internal helper to resolve a SharePoint or OneDrive URL to get the IDs needed for Excel operations.
    
    Args:
        url: The SharePoint or OneDrive URL pointing to a site or document library.
        file_name: The name of the Excel file (with .xlsx extension).
    
    Returns:
        Dictionary with status, site_id, drive_id, item_id, and file metadata.
        On error, returns dict with status='error' and message.
    """
    try:
        headers = await get_graph_headers()
        
        # Parse the URL
        parsed = parse_sharepoint_url(url)
        hostname = parsed["hostname"]
        site_path = parsed["site_path"]
        
        if not hostname:
            return {
                "status": "error",
                "message": "Could not parse hostname from URL",
            }
        
        logger.info(f"Resolving URL - hostname: {hostname}, site_path: {site_path}, file_name: {file_name}")
        
        async with httpx.AsyncClient() as client:
            # Step 1: Get the site ID
            if site_path:
                site_url = f"{GRAPH_API_BASE}/sites/{hostname}:{site_path}"
            else:
                site_url = f"{GRAPH_API_BASE}/sites/{hostname}"
            
            logger.info(f"Getting site from: {site_url}")
            site_response = await client.get(site_url, headers=headers, timeout=30.0)
            
            if site_response.status_code != 200:
                error_data = site_response.json() if site_response.content else {}
                error_message = error_data.get("error", {}).get("message", site_response.text)
                return {
                    "status": "error",
                    "message": f"Failed to get site: {error_message}",
                    "status_code": site_response.status_code,
                }
            
            site_data = site_response.json()
            site_id = site_data["id"]
            logger.info(f"Found site ID: {site_id}")
            
            # Step 2: Get the drives (document libraries) for the site
            drives_url = f"{GRAPH_API_BASE}/sites/{site_id}/drives"
            drives_response = await client.get(drives_url, headers=headers, timeout=30.0)
            
            if drives_response.status_code != 200:
                error_data = drives_response.json() if drives_response.content else {}
                error_message = error_data.get("error", {}).get("message", drives_response.text)
                return {
                    "status": "error",
                    "message": f"Failed to get drives: {error_message}",
                    "status_code": drives_response.status_code,
                }
            
            drives_data = drives_response.json()
            
            # Find the Documents/Shared Documents drive
            drive_id = None
            drive_name = None
            for drive in drives_data.get("value", []):
                # Look for the default document library
                if drive.get("name") in ("Documents", "Shared Documents") or \
                   drive.get("driveType") == "documentLibrary":
                    drive_id = drive["id"]
                    drive_name = drive.get("name")
                    break
            
            if not drive_id and drives_data.get("value"):
                # Fall back to the first drive
                drive_id = drives_data["value"][0]["id"]
                drive_name = drives_data["value"][0].get("name")
            
            if not drive_id:
                return {
                    "status": "error",
                    "message": "No document library found for this site",
                }
            
            logger.info(f"Found drive ID: {drive_id} ({drive_name})")
            
            # Step 3: Get the file item by path
            item_url = f"{GRAPH_API_BASE}/drives/{drive_id}/root:/{file_name}"
            logger.info(f"Getting item from: {item_url}")
            item_response = await client.get(item_url, headers=headers, timeout=30.0)
            
            if item_response.status_code != 200:
                error_data = item_response.json() if item_response.content else {}
                error_message = error_data.get("error", {}).get("message", item_response.text)
                return {
                    "status": "error", 
                    "message": f"Failed to get item: {error_message}",
                    "status_code": item_response.status_code,
                    "site_id": site_id,
                    "drive_id": drive_id,
                    "attempted_path": file_name,
                }
            
            item_data = item_response.json()
            item_id = item_data["id"]
            
            logger.info(f"Found item ID: {item_id}")
            
            # Return all the resolved IDs
            return {
                "status": "success",
                "message": "Successfully resolved all IDs from URL",
                "site_id": site_id,
                "site_name": site_data.get("displayName"),
                "drive_id": drive_id,
                "drive_name": drive_name,
                "item_id": item_id,
                "file_name": item_data.get("name"),
                "file_path": file_name,
                "web_url": item_data.get("webUrl"),
                "size": item_data.get("size"),
                "last_modified": item_data.get("lastModifiedDateTime"),
            }
            
    except httpx.HTTPError as e:
        logger.error(f"HTTP error resolving URL: {e}")
        return {
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }
    except Exception as e:
        logger.error(f"Error resolving URL: {e}")
        return {
            "status": "error",
            "message": str(e),
        }


# =============================================================================
# Core Implementation Functions (called by tools)
# =============================================================================

async def _update_row_by_lookup_impl(
    url: str,
    file_name: str,
    sheet_name: str,
    search_column: str,
    reference_value: str,
    target_columns_list: list,
    values_list: list,
    row_offset: int = 0
) -> dict:
    """
    Core implementation for updating a row by lookup.
    Returns a dict with the result (not a JSON string).
    This is the internal implementation called by both tools.
    """
    # Validate that columns and values have the same length
    if len(target_columns_list) != len(values_list):
        return {
            "status": "error",
            "message": f"Mismatch: {len(target_columns_list)} columns but {len(values_list)} values provided. They must be equal.",
        }
    
    # Resolve URL to get drive_id, item_id, and site_id
    resolved = await resolve_excel_file_ids(url, file_name)
    if resolved.get("status") != "success":
        return resolved
    
    drive_id = resolved["drive_id"]
    item_id = resolved["item_id"]
    site_id = resolved.get("site_id")
    
    workbook_url = build_workbook_url(drive_id, item_id, site_id)
    headers = await get_graph_headers()
    
    async with httpx.AsyncClient() as client:
        # Step 1: Get the used range to find the data extent
        used_range_url = f"{workbook_url}/worksheets/{sheet_name}/usedRange"
        logger.info(f"Getting used range for sheet '{sheet_name}'")
        
        used_range_response = await client.get(
            used_range_url,
            headers=headers,
            timeout=30.0,
        )
        
        if used_range_response.status_code != 200:
            error_data = used_range_response.json() if used_range_response.content else {}
            error_message = error_data.get("error", {}).get("message", used_range_response.text)
            return {
                "status": "error",
                "message": f"Failed to get used range: {error_message}",
                "status_code": used_range_response.status_code,
            }
        
        used_range_data = used_range_response.json()
        row_count = used_range_data.get("rowCount", 0)
        
        if row_count == 0:
            return {
                "status": "error",
                "message": "Worksheet is empty",
            }
        
        # Step 2: Get the search column values
        search_range = f"{search_column}1:{search_column}{row_count}"
        search_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{search_range}')"
        logger.info(f"Searching column {search_column} for value '{reference_value}'")
        
        search_response = await client.get(
            search_url,
            headers=headers,
            timeout=30.0,
        )
        
        if search_response.status_code != 200:
            error_data = search_response.json() if search_response.content else {}
            error_message = error_data.get("error", {}).get("message", search_response.text)
            return {
                "status": "error",
                "message": f"Failed to read search column: {error_message}",
                "status_code": search_response.status_code,
            }
        
        search_data = search_response.json()
        column_values = search_data.get("values", [])
        
        # Step 3: Find the row with the reference value
        found_row = None
        for i, row in enumerate(column_values):
            cell_value = row[0] if row else None
            # Use smart comparison that handles date conversions
            if compare_values_for_search(cell_value, reference_value):
                found_row = i + 1  # Excel rows are 1-indexed
                logger.info(f"Found match: cell value '{cell_value}' matches reference '{reference_value}'")
                break
        
        if found_row is None:
            # Provide more diagnostic info in the error message
            sample_values = [str(row[0]) if row and row[0] is not None else "empty" 
                            for row in column_values[:10]]
            return {
                "status": "error",
                "message": f"Reference value '{reference_value}' not found in column {search_column}",
                "searched_rows": row_count,
                "sample_values": sample_values,
                "hint": "If searching for a date, ensure format matches (e.g., '12/22/2025' or '2025-12-22')"
            }
        
        # Apply row offset
        target_row = found_row + row_offset
        logger.info(f"Found reference value in row {found_row}, target row is {target_row} (offset: {row_offset})")
        
        # Step 4: Update each cell individually
        updated_cells = []
        errors = []
        
        for col, value in zip(target_columns_list, values_list):
            cell_address = f"{col.upper()}{target_row}"
            cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{cell_address}')"
            
            body = {
                "values": [[value]],
            }
            
            logger.info(f"Updating cell '{cell_address}' with value '{value}'")
            
            update_response = await client.patch(
                cell_url,
                headers=headers,
                json=body,
                timeout=30.0,
            )
            
            if update_response.status_code == 200:
                updated_cells.append(cell_address)
            else:
                error_data = update_response.json() if update_response.content else {}
                error_message = error_data.get("error", {}).get("message", update_response.text)
                errors.append({"cell": cell_address, "error": error_message})
        
        if errors:
            return {
                "status": "partial_error",
                "message": f"Some cells failed to update",
                "updated_cells": updated_cells,
                "errors": errors,
            }
        
        return {
            "status": "success",
            "message": f"Successfully updated {len(values_list)} cells in row {target_row}",
            "sheet_name": sheet_name,
            "found_row": found_row,
            "target_row": target_row,
            "row_offset": row_offset,
            "reference_value": reference_value,
            "updated_cells": updated_cells,
            "columns": target_columns_list,
            "values_written": len(values_list),
        }


# =============================================================================
# MCP Tools
# =============================================================================

@mcp.tool(name="excel.updateRowByLookup")
async def excel_update_row_by_lookup(
    url: str,
    file_name: str,
    sheet_name: str,
    search_column: str,
    reference_value: str,
    target_columns: str,
    values: str,
    row_offset: int = 0
) -> str:
    """
    Find a row by looking up a reference value and update specific columns in that row.
    
    This tool searches for a specific value in a column, finds the row containing that value,
    and then updates the specified columns with the provided values. Use row_offset to
    update a row below the found row (e.g., row_offset=1 updates the next row).
    
    Args:
        url: SharePoint/OneDrive URL to the document library (e.g., https://contoso.sharepoint.com/sites/Sales/Shared%20Documents)
        file_name: Excel file name with .xlsx extension (e.g., "Budget.xlsx")
        sheet_name: Worksheet name (e.g., "Sheet1")
        search_column: Column letter to search (e.g., "A" or "C")
        reference_value: Value to find in the search column (supports dates like "12/22/2025")
        target_columns: JSON array of column letters to update (e.g., '["D", "F", "H"]')
        values: JSON array of values to write, must match length of target_columns (e.g., '["value1", 123, true]')
        row_offset: Number of rows below the found row to update. Default 0 = same row, 1 = next row
    
    Returns:
        JSON string with operation result
    """
    try:
        # Parse JSON string parameters into lists
        try:
            target_columns_list = json.loads(target_columns)
            if not isinstance(target_columns_list, list):
                return json.dumps({
                    "status": "error",
                    "message": "target_columns must be a JSON array of column letters (e.g., '[\"D\", \"F\", \"H\"]')",
                }, indent=2)
        except json.JSONDecodeError as e:
            return json.dumps({
                "status": "error",
                "message": f"Invalid JSON in target_columns: {str(e)}. Expected format: '[\"D\", \"F\", \"H\"]'",
            }, indent=2)
        
        try:
            values_list = json.loads(values)
            if not isinstance(values_list, list):
                return json.dumps({
                    "status": "error",
                    "message": "values must be a JSON array (e.g., '[\"value1\", 123, true]')",
                }, indent=2)
        except json.JSONDecodeError as e:
            return json.dumps({
                "status": "error",
                "message": f"Invalid JSON in values: {str(e)}. Expected format: '[\"value1\", 123, true]'",
            }, indent=2)
        
        # Call the core implementation
        result = await _update_row_by_lookup_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            search_column=search_column,
            reference_value=reference_value,
            target_columns_list=target_columns_list,
            values_list=values_list,
            row_offset=row_offset
        )
        
        return json.dumps(result, indent=2)
                
    except httpx.HTTPError as e:
        logger.error(f"HTTP error in updateRowByLookup: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error in updateRowByLookup: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


@mcp.tool(name="excel.updateRange")
async def excel_update_range(
    url: str,
    file_name: str,
    sheet_name: str,
    address: str,
    values: str
) -> str:
    """
    Update a range of cells in an Excel worksheet using Microsoft Graph API.
    
    Args:
        url: SharePoint/OneDrive URL to the document library (e.g., https://contoso.sharepoint.com/sites/Sales/Shared%20Documents)
        file_name: Excel file name with .xlsx extension (e.g., "Budget.xlsx")
        sheet_name: Worksheet name (e.g., "Sheet1")
        address: Cell range address (e.g., "A1:C3", "B2:D10")
        values: JSON 2D array of values, where each inner array is a row (e.g., '[["row1col1", "row1col2"], ["row2col1", "row2col2"]]')
    
    Returns:
        JSON string with operation result
    """
    try:
        # Parse JSON string parameter into 2D list
        try:
            values_list = json.loads(values)
            if not isinstance(values_list, list) or not all(isinstance(row, list) for row in values_list):
                return json.dumps({
                    "status": "error",
                    "message": "values must be a JSON 2D array (e.g., '[[\"a\", \"b\"], [\"c\", \"d\"]]')",
                }, indent=2)
        except json.JSONDecodeError as e:
            return json.dumps({
                "status": "error",
                "message": f"Invalid JSON in values: {str(e)}. Expected format: '[[\"a\", \"b\"], [\"c\", \"d\"]]'",
            }, indent=2)
        
        # Resolve URL to get drive_id, item_id, and site_id
        resolved = await resolve_excel_file_ids(url, file_name)
        if resolved.get("status") != "success":
            return json.dumps(resolved, indent=2)
        
        drive_id = resolved["drive_id"]
        item_id = resolved["item_id"]
        site_id = resolved.get("site_id")
        
        # Build the URL for range update
        workbook_url = build_workbook_url(drive_id, item_id, site_id)
        # URL encode the sheet name and address for special characters
        range_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"
        
        # Build request body
        body = {
            "values": values_list,
        }
        
        headers = await get_graph_headers()
        
        logger.info(f"Updating range '{address}' in sheet '{sheet_name}'")
        
        async with httpx.AsyncClient() as client:
            response = await client.patch(
                range_url,
                headers=headers,
                json=body,
                timeout=30.0,
            )
            
            if response.status_code == 200:
                result_data = response.json()
                result = {
                    "status": "success",
                    "message": f"Successfully updated range '{address}' in sheet '{sheet_name}'",
                    "file_name": resolved.get("file_name"),
                    "sheet_name": sheet_name,
                    "address": result_data.get("address", address),
                    "row_count": result_data.get("rowCount"),
                    "column_count": result_data.get("columnCount"),
                }
                logger.info(f"Successfully updated range '{address}' in sheet '{sheet_name}'")
                return json.dumps(result, indent=2)
            else:
                error_data = response.json() if response.content else {}
                error_message = error_data.get("error", {}).get("message", response.text)
                return json.dumps({
                    "status": "error",
                    "message": f"Failed to update range: {error_message}",
                    "status_code": response.status_code,
                }, indent=2)
                
    except httpx.HTTPError as e:
        logger.error(f"HTTP error updating range: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error updating range: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


@mcp.tool(name="excel.logTrades")
async def excel_log_trades(
    trades: str,
    reference_date: str = "",
    sheet_name: str = ""
) -> str:
    """
    Log multiple trades to the configured trade tracker spreadsheet.
    
    This tool appends trade data below a reference date row in the spreadsheet.
    The spreadsheet URL and file name are configured via environment variables
    (TRADE_TRACKER_URL, TRADE_TRACKER_FILE).
    
    Args:
        trades: JSON array of trade objects. Each object can have:
                - open_date: Date when trade was opened (e.g., "12/23/2025")
                - open_time: Time when trade was opened (e.g., "10:30 AM")
                - close_date: Date when trade was closed, if known (e.g., "12/27/2025")
                - close_time: Time when trade was closed, if known (e.g., "4:00 PM")
                - strategy: Strategy name (e.g., "VPCS", "IC", "Iron Condor")
                - credit: Credit received when opening (number, e.g., 0.25)
                - debit: Debit paid if closed before expiration (number, e.g., 0.10)
                - contracts: Number of contracts (integer, e.g., 25)
                - open_fees: Total fees paid when trade was opened (number, e.g., 176.58)
                - close_fees: Total fees paid if closed before expiration (number, e.g., 88.29)
                - sold_call_strike: Strike price for sold calls (number, e.g., 6100)
                - sold_put_strike: Strike price for sold puts (number, e.g., 5800)
                - width: Width in USD between sold and bought strikes (number, e.g., 15)
                - expired: Boolean flag indicating if the trade expired (default: false).
                          When true (or when debit=0), auto-fills:
                          - close_date = open_date (0DTE expires same day)
                          - close_time = "4:00 PM" (market close)
                          - debit = 0 (expired worthless)
                
                Example: '[{"open_date": "12/23/2025", "open_time": "10:30 AM", "strategy": "IC", 
                           "credit": 0.60, "contracts": 25, "open_fees": 176.58, 
                           "sold_call_strike": 6100, "sold_put_strike": 5800, "width": 15, "expired": true}]'
        reference_date: Date to search for in column C. Trades will be appended below this row.
                       Format: "MM/DD/YYYY" (e.g., "12/22/2025"). If not provided, uses the
                       last non-empty date value in column C.
        sheet_name: Worksheet name (default: current month, e.g., "December").
    
    Returns:
        JSON string with operation result including count of trades logged.
    
    Column Mapping:
        - Column C: Date when the trade was opened
        - Column E: Time when the trade was opened
        - Column F: Date when the trade was closed (if known)
        - Column G: Time when the trade was closed (if known, usually 4:00 PM)
        - Column I: Strategy
        - Column J: Credit Received
        - Column K: Debit Paid (only when trade is closed before expiration)
        - Column L: Number of Contracts
        - Column N: Total fees paid when the trade was opened
        - Column O: Total fees paid if the trade was closed before expiration
        - Column Q: Strike price for Sold Calls
        - Column R: Strike price for Sold Puts
        - Column T: Width in USD between sold and bought strikes
    
    Expired Trade Handling:
        For 0DTE SPX options that expire worthless:
        - Set "expired": true in the trade object, OR
        - Set "debit": 0 (tool will auto-detect as expired)
        The tool will automatically populate close_date with open_date
        and close_time with "4:00 PM" (market close).
    """
    # Default sheet_name to current month if not provided
    if not sheet_name:
        sheet_name = datetime.now().strftime("%B")  # e.g., "December"
    
    # Column configuration for trade tracker
    # C=open_date, E=open_time, F=close_date, G=close_time, I=strategy, J=credit, 
    # K=debit, L=contracts, N=open_fees, O=close_fees, Q=sold_call, R=sold_put, T=width
    COLUMN_ORDER = ["C", "E", "F", "G", "I", "J", "K", "L", "N", "O", "Q", "R", "T"]
    
    # Log incoming parameters for debugging
    logger.info(f"excel.logTrades called with reference_date='{reference_date}', sheet_name='{sheet_name}'")
    logger.info(f"Raw trades input (first 500 chars): {trades[:500] if len(trades) > 500 else trades}")
    
    try:
        # Parse the trades JSON
        try:
            trades_list = json.loads(trades)
            logger.info(f"Parsed {len(trades_list)} trades from JSON")
            if not isinstance(trades_list, list):
                return json.dumps({
                    "status": "error",
                    "message": "trades must be a JSON array of trade objects",
                }, indent=2)
        except json.JSONDecodeError as e:
            return json.dumps({
                "status": "error",
                "message": f"Invalid JSON in trades: {str(e)}",
            }, indent=2)
        
        if len(trades_list) == 0:
            return json.dumps({
                "status": "warning",
                "message": "No trades provided to log",
            }, indent=2)
        
        # Sort trades by open_date (ascending) then by open_time (ascending)
        def parse_trade_datetime(trade: dict) -> tuple:
            """
            Parse trade's open_date and open_time for sorting.
            Returns a tuple (date_key, time_key) for proper chronological ordering.
            """
            open_date_str = trade.get("open_date", trade.get("date", ""))
            open_time_str = trade.get("open_time", trade.get("time", ""))
            
            # Parse date - default to max date if invalid
            date_key = datetime.max
            if open_date_str:
                parsed_date = parse_date_string(open_date_str)
                if parsed_date:
                    date_key = parsed_date
            
            # Parse time - default to max time if invalid
            time_key = datetime.max.time()
            if open_time_str:
                # Try common time formats
                time_formats = [
                    "%I:%M %p",      # 10:30 AM
                    "%I:%M%p",       # 10:30AM
                    "%H:%M",         # 14:30
                    "%H:%M:%S",      # 14:30:00
                    "%I:%M:%S %p",   # 10:30:00 AM
                ]
                for fmt in time_formats:
                    try:
                        parsed_time = datetime.strptime(open_time_str.strip().upper(), fmt)
                        time_key = parsed_time.time()
                        break
                    except ValueError:
                        continue
            
            return (date_key, time_key)
        
        # Sort trades chronologically (earliest first)
        trades_list.sort(key=parse_trade_datetime)
        logger.info(f"Sorted {len(trades_list)} trades by open_date and open_time (ascending)")
        
        # If reference_date not provided, find the last non-empty date in column C
        if not reference_date:
            logger.info(f"No reference_date provided, finding last date in column C of sheet '{sheet_name}'")
            
            # Resolve URL to get drive_id, item_id, and site_id
            resolved = await resolve_excel_file_ids(TRADE_TRACKER_URL, TRADE_TRACKER_FILE)
            if resolved.get("status") != "success":
                return json.dumps({
                    "status": "error",
                    "message": f"Failed to resolve Excel file: {resolved.get('message')}",
                }, indent=2)
            
            drive_id = resolved["drive_id"]
            item_id = resolved["item_id"]
            site_id = resolved.get("site_id")
            workbook_url = build_workbook_url(drive_id, item_id, site_id)
            headers = await get_graph_headers()
            
            async with httpx.AsyncClient() as client:
                # Get used range to find data extent
                used_range_url = f"{workbook_url}/worksheets/{sheet_name}/usedRange"
                used_range_response = await client.get(used_range_url, headers=headers, timeout=30.0)
                
                if used_range_response.status_code != 200:
                    error_data = used_range_response.json() if used_range_response.content else {}
                    error_message = error_data.get("error", {}).get("message", used_range_response.text)
                    return json.dumps({
                        "status": "error",
                        "message": f"Failed to get worksheet data: {error_message}",
                    }, indent=2)
                
                used_range_data = used_range_response.json()
                row_count = used_range_data.get("rowCount", 0)
                
                if row_count == 0:
                    return json.dumps({
                        "status": "error",
                        "message": f"Worksheet '{sheet_name}' is empty",
                    }, indent=2)
                
                # Get column C values
                search_range = f"C1:C{row_count}"
                search_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{search_range}')"
                search_response = await client.get(search_url, headers=headers, timeout=30.0)
                
                if search_response.status_code != 200:
                    error_data = search_response.json() if search_response.content else {}
                    error_message = error_data.get("error", {}).get("message", search_response.text)
                    return json.dumps({
                        "status": "error",
                        "message": f"Failed to read column C: {error_message}",
                    }, indent=2)
                
                search_data = search_response.json()
                column_values = search_data.get("values", [])
                
                # Find the last non-empty cell with a valid date value (searching from bottom)
                last_date_value = None
                last_date_row = None
                
                for i in range(len(column_values) - 1, -1, -1):
                    cell_value = column_values[i][0] if column_values[i] else None
                    if cell_value is not None and cell_value != "":
                        # Check if it's a valid date (either Excel serial number or date string)
                        try:
                            # If it's a number (Excel serial date), convert to date string
                            if isinstance(cell_value, (int, float)) and 1 <= cell_value <= 2958465:
                                dt = excel_serial_to_date(cell_value)
                                reference_date = dt.strftime("%m/%d/%Y")
                                last_date_value = cell_value
                                last_date_row = i + 1
                                logger.info(f"Found last date at row {last_date_row}: Excel serial {cell_value} → {reference_date}")
                                break
                            # If it's a string, try to parse it as a date
                            elif isinstance(cell_value, str):
                                parsed = parse_date_string(cell_value)
                                if parsed:
                                    reference_date = cell_value
                                    last_date_value = cell_value
                                    last_date_row = i + 1
                                    logger.info(f"Found last date at row {last_date_row}: {reference_date}")
                                    break
                        except (ValueError, TypeError):
                            continue
                
                if not reference_date:
                    return json.dumps({
                        "status": "error",
                        "message": f"Could not find any valid date in column C of sheet '{sheet_name}'",
                    }, indent=2)
        
        logger.info(f"Logging {len(trades_list)} trades to {TRADE_TRACKER_FILE}, sheet '{sheet_name}'")
        logger.info(f"Reference date: {reference_date}, URL: {TRADE_TRACKER_URL}")
        
        results = []
        errors = []
        
        for i, trade in enumerate(trades_list):
            # Log the raw trade object for debugging
            logger.info(f"Processing trade {i+1}: {json.dumps(trade)}")
            
            # Extract trade fields with defaults
            # Map strategy name to Excel short code
            raw_strategy = trade.get("strategy", "")
            mapped_strategy = map_strategy_name(raw_strategy)
            if raw_strategy != mapped_strategy:
                logger.info(f"Mapped strategy '{raw_strategy}' → '{mapped_strategy}'")
            
            # Get open date/time
            open_date = trade.get("open_date", trade.get("date", ""))
            open_time = trade.get("open_time", trade.get("time", ""))
            
            # Handle expired trades
            # If expired=true, or if debit=0 and no close_date provided, treat as expired
            expired_raw = trade.get("expired", False)
            # Handle string "true"/"false" from some AI agents that serialize booleans as strings
            if isinstance(expired_raw, str):
                is_expired = expired_raw.lower() in ("true", "1", "yes")
            else:
                is_expired = bool(expired_raw)
            
            close_date = trade.get("close_date", "")
            close_time = trade.get("close_time", "")
            debit = trade.get("debit", "")
            
            logger.info(f"Trade {i+1}: expired_raw={expired_raw!r} (type={type(expired_raw).__name__}), is_expired={is_expired}, close_date='{close_date}', debit={debit}")
            
            # Auto-detect expiration: if debit is 0 (or not provided) and no close info
            if not close_date:
                # Check if explicitly marked as expired
                if is_expired:
                    close_date = open_date  # Expired on the same day
                    close_time = close_time or "4:00 PM"  # Market close
                    debit = 0 if debit == "" else debit  # Expired worthless = $0 debit
                    logger.info(f"Trade marked as expired: close_date={close_date}, close_time={close_time}, debit={debit}")
                # Or if debit is explicitly 0, also treat as expired
                elif debit == 0:
                    close_date = open_date
                    close_time = close_time or "4:00 PM"
                    logger.info(f"Trade with debit=0 treated as expired: close_date={close_date}, close_time={close_time}")
            
            # Build values array matching COLUMN_ORDER
            # Use empty string for optional fields that aren't provided
            values = [
                open_date,                                           # C: Open date
                open_time,                                           # E: Open time
                close_date,                                          # F: Close date
                close_time,                                          # G: Close time
                mapped_strategy,                                      # I: Strategy
                trade.get("credit", ""),                             # J: Credit received
                debit,                                               # K: Debit paid (if closed early)
                trade.get("contracts", ""),                          # L: Number of contracts
                trade.get("open_fees", trade.get("fees", "")),       # N: Open fees (backward compat with "fees")
                trade.get("close_fees", ""),                         # O: Close fees
                trade.get("sold_call_strike", ""),                   # Q: Sold call strike
                trade.get("sold_put_strike", ""),                    # R: Sold put strike
                trade.get("width", ""),                              # T: Width between strikes
            ]
            
            # Log the values array for debugging
            logger.info(f"Trade {i+1} values array: C={values[0]}, E={values[1]}, F={values[2]}, G={values[3]}, I={values[4]}")
            
            # Calculate row offset: first trade goes to row_offset=1, second to row_offset=2, etc.
            current_offset = 1 + i
            
            logger.info(f"Logging trade {i+1}/{len(trades_list)}: strategy={mapped_strategy}, credit={trade.get('credit')}, contracts={trade.get('contracts')} at offset {current_offset}")
            
            # Call the core implementation directly (not the decorated tool)
            result_data = await _update_row_by_lookup_impl(
                url=TRADE_TRACKER_URL,
                file_name=TRADE_TRACKER_FILE,
                sheet_name=sheet_name,
                search_column="C",
                reference_value=reference_date,
                target_columns_list=COLUMN_ORDER,
                values_list=values,
                row_offset=current_offset
            )
            
            if result_data.get("status") == "success":
                results.append({
                    "trade_index": i + 1,
                    "row": result_data.get("target_row"),
                    "open_date": open_date,
                    "open_time": open_time,
                    "close_date": close_date,
                    "close_time": close_time,
                    "strategy": mapped_strategy,
                    "credit": trade.get("credit", ""),
                    "debit": debit,
                    "contracts": trade.get("contracts", ""),
                    "open_fees": trade.get("open_fees", trade.get("fees", "")),
                    "close_fees": trade.get("close_fees", ""),
                    "sold_call_strike": trade.get("sold_call_strike", ""),
                    "sold_put_strike": trade.get("sold_put_strike", ""),
                    "width": trade.get("width", ""),
                    "expired": is_expired,
                })
            else:
                errors.append({
                    "trade_index": i + 1,
                    "error": result_data.get("message"),
                    "strategy": mapped_strategy,
                })
        
        # Build response
        if errors and not results:
            return json.dumps({
                "status": "error",
                "message": f"All {len(trades_list)} trades failed to log",
                "errors": errors,
            }, indent=2)
        elif errors:
            return json.dumps({
                "status": "partial_success",
                "message": f"Logged {len(results)} of {len(trades_list)} trades",
                "trades_logged": len(results),
                "trades_failed": len(errors),
                "file_name": TRADE_TRACKER_FILE,
                "sheet_name": sheet_name,
                "reference_date": reference_date,
                "results": results,
                "errors": errors,
            }, indent=2)
        else:
            return json.dumps({
                "status": "success",
                "message": f"Successfully logged {len(results)} trades",
                "trades_logged": len(results),
                "file_name": TRADE_TRACKER_FILE,
                "sheet_name": sheet_name,
                "reference_date": reference_date,
                "results": results,
            }, indent=2)
            
    except Exception as e:
        logger.error(f"Error logging trades: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


# Health check endpoint (for Container Apps)
@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Health check endpoint for Container Apps."""
    return JSONResponse({"status": "healthy", "service": "mcp-excel-server"})


if __name__ == "__main__":
    port = int(os.getenv("PORT", "3000"))
    host = os.getenv("HOST", "0.0.0.0")
    
    logger.info(f"Starting MCP Excel Service on {host}:{port}")
    logger.info(f"MCP endpoint: http://{host}:{port}/mcp")
    
    # Run with HTTP transport (streamable HTTP)
    mcp.run(transport="http", host=host, port=port)
