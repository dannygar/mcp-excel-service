"""
Excel Helper Functions Module

Provides utility functions for working with Excel files:
- Date parsing and conversion (string <-> Excel serial numbers)
- SharePoint/OneDrive URL parsing
- Excel file ID resolution via Microsoft Graph API
"""

import re
import logging
from datetime import datetime, timedelta
from typing import Optional, Union
from urllib.parse import urlparse, unquote

import httpx

from graph_api import GRAPH_API_BASE, get_graph_headers

logger = logging.getLogger("mcp-excel-server")


# =============================================================================
# Date Handling Functions
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


# =============================================================================
# URL Parsing Functions
# =============================================================================

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


# =============================================================================
# Excel File Resolution
# =============================================================================

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
