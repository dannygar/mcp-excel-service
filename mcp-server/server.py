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
from typing import Optional

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mcp-excel-server")

# Load environment variables
load_dotenv()

# Initialize MCP server
mcp = FastMCP(
    "MCP Excel Service",
    dependencies=["httpx", "python-dotenv"],
)

# Microsoft Graph API base URL
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# Token cache for client credentials flow
_token_cache = {
    "access_token": None,
    "expires_at": 0,
}


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


@mcp.tool(name="excel.appendRows")
async def excel_append_rows(
    drive_id: str,
    item_id: str,
    table_name: str,
    rows: list[list],
    site_id: Optional[str] = None,
    persist_changes: bool = True,
) -> str:
    """
    Append rows to an Excel table using Microsoft Graph API.
    
    Args:
        drive_id: The ID of the drive containing the Excel file
        item_id: The ID of the Excel file item
        table_name: The name of the table to append rows to
        rows: A 2D array of values to append (each inner array is a row)
        site_id: Optional SharePoint site ID (for SharePoint-hosted files)
        persist_changes: Whether to persist changes immediately (default: True)
    
    Returns:
        JSON string with the result of the operation
    
    Example:
        excel_append_rows(
            drive_id="b!abc123",
            item_id="xyz789",
            table_name="SalesData",
            rows=[["Product A", 100, 25.99], ["Product B", 50, 15.99]]
        )
    """
    try:
        # Build the URL for table row append
        workbook_url = build_workbook_url(drive_id, item_id, site_id)
        url = f"{workbook_url}/tables/{table_name}/rows"
        
        # Build request body
        body = {
            "values": rows,
        }
        
        # Add session header if persist_changes is specified
        headers = await get_graph_headers()
        if not persist_changes:
            headers["Workbook-Session-Id"] = "non-persistent"
        
        logger.info(f"Appending {len(rows)} rows to table '{table_name}'")
        
        async with httpx.AsyncClient() as client:
            response = await client.post(
                url,
                headers=headers,
                json=body,
                timeout=30.0,
            )
            
            if response.status_code == 201:
                result_data = response.json()
                result = {
                    "status": "success",
                    "message": f"Successfully appended {len(rows)} rows to table '{table_name}'",
                    "table_name": table_name,
                    "rows_added": len(rows),
                    "row_index": result_data.get("index"),
                }
                logger.info(f"Successfully appended rows to table '{table_name}'")
                return json.dumps(result, indent=2)
            else:
                error_data = response.json() if response.content else {}
                error_message = error_data.get("error", {}).get("message", response.text)
                return json.dumps({
                    "status": "error",
                    "message": f"Failed to append rows: {error_message}",
                    "status_code": response.status_code,
                }, indent=2)
                
    except httpx.HTTPError as e:
        logger.error(f"HTTP error appending rows: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error appending rows: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


@mcp.tool(name="excel.updateRange")
async def excel_update_range(
    drive_id: str,
    item_id: str,
    sheet_name: str,
    address: str,
    values: list[list],
) -> str:
    """
    Update a range of cells in an Excel worksheet using Microsoft Graph API.
    
    Args:
        drive_id: The ID of the drive containing the Excel file
        item_id: The ID of the Excel file item
        sheet_name: The name of the worksheet
        address: The cell range address (e.g., "A1:C3", "B2:D10")
        values: A 2D array of values to set in the range
    
    Returns:
        JSON string with the result of the operation
    
    Example:
        excel_update_range(
            drive_id="b!abc123",
            item_id="xyz789",
            sheet_name="Sheet1",
            address="A1:C2",
            values=[["Name", "Quantity", "Price"], ["Widget", 100, 9.99]]
        )
    """
    try:
        # Build the URL for range update
        workbook_url = build_workbook_url(drive_id, item_id)
        # URL encode the sheet name and address for special characters
        url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"
        
        # Build request body
        body = {
            "values": values,
        }
        
        headers = await get_graph_headers()
        
        logger.info(f"Updating range '{address}' in sheet '{sheet_name}'")
        
        async with httpx.AsyncClient() as client:
            response = await client.patch(
                url,
                headers=headers,
                json=body,
                timeout=30.0,
            )
            
            if response.status_code == 200:
                result_data = response.json()
                result = {
                    "status": "success",
                    "message": f"Successfully updated range '{address}' in sheet '{sheet_name}'",
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
