"""
MCP Excel Service - Streamable HTTP Transport

This server provides Excel manipulation tools via the Model Context Protocol (MCP).
Uses Microsoft Graph API to interact with Excel files in SharePoint/OneDrive.
Designed for deployment on Azure Container Apps with Foundry Agent integration.

Authentication:
    - Incoming requests: Validates Microsoft Entra ID bearer tokens from Foundry agents.
      Requires Project Managed Identity authentication with valid audience.
    - Outgoing Graph API calls: Uses Azure AD service principal (client credentials flow).
    
    Required environment variables:
    - AZURE_TENANT_ID: Azure AD tenant ID
    - AZURE_CLIENT_ID: App registration client ID (also used as audience for token validation)
    - AZURE_CLIENT_SECRET: App registration client secret (for Graph API access)

Security:
    - All MCP endpoints require valid Entra ID bearer token (401 for invalid/missing)
    - Token audience must match AZURE_CLIENT_ID
    - Token issuer must be from configured tenant
    - Health endpoint (/health) is public for Container Apps probes

REST API:
    - In addition to MCP protocol, the server exposes REST API endpoints at /api/v1/*
    - OpenAPI specification available at /api/v1/openapi.json
    - Same authentication (Entra ID bearer tokens) required for REST endpoints
"""

import os
import json
import logging
import pathlib
from datetime import datetime

from pydantic import BaseModel, Field

from dotenv import load_dotenv

# =============================================================================
# Load environment variables FIRST (before any other imports that use env vars)
# =============================================================================
# Use MCP_ENV to select environment:
#   - "local" (default): Uses config/.env.local for local development
#   - "dev": Uses config/.env.dev for Azure dev deployment
#   - "prod": Uses config/.env.prod for Azure production deployment

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mcp-excel-server")

_project_root = pathlib.Path(__file__).parent.parent
_mcp_env = os.getenv("MCP_ENV", "local")  # Default to 'local' for local development
_env_file = _project_root / "config" / f".env.{_mcp_env}"

if _env_file.exists():
    load_dotenv(_env_file)
    logger.info(f"Loaded config from {_env_file}")
else:
    # Fallback to .env.local if specified env doesn't exist
    _env_local = _project_root / "config" / ".env.local"
    if _env_local.exists():
        load_dotenv(_env_local)
        logger.info(f"Loaded config from {_env_local} (fallback)")
    else:
        load_dotenv()  # Fall back to default behavior
        logger.warning("No .env file found, using environment variables only")

# =============================================================================
# Now import modules that depend on environment variables
# =============================================================================

import httpx
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse

# Import authentication components
from auth import (
    ENTRA_AUTH_ENABLED,
    ENTRA_TENANT_ID,
    ENTRA_CLIENT_ID,
    EntraAuthMiddleware,
    configure_auth_middleware,
    create_validator_from_env,
)

# Import Graph API helpers
from graph_api import (
    get_graph_headers,
    build_workbook_url,
)

# Import Excel helpers
from excel_helpers import (
    parse_date_string,
    excel_serial_to_date,
    resolve_excel_file_ids,
)

# Import core operations
from core_operations import update_trade_with_delta_impl, close_trade_impl, update_range_impl, log_trades_impl

# Import configuration
from config import map_strategy_name

# Initialize MCP server
mcp = FastMCP("MCP Excel Service")


# =============================================================================
# MCP Tools
# =============================================================================

@mcp.tool(name="excel.updateTradeWithDelta")
async def excel_update_trade_with_delta(
    url: str,
    file_name: str,
    sheet_name: str,
    trade_date: str,
    trade_time: str,
    sold_strike: str,
    delta: float,
    delta_time: str
) -> str:
    """
    Update delta value for a trade in an Excel workbook.
    
    This tool finds a trade row by matching the trade date (Column A) and trade time (Column C),
    then updates the appropriate delta column based on the delta time and strike type.
    
    Args:
        url: SharePoint/OneDrive URL to the document library (e.g., https://contoso.sharepoint.com/Shared%20Documents)
        file_name: Excel workbook name with .xlsx extension (e.g., "2026 Trade Tracker.xlsx")
        sheet_name: Worksheet name (e.g., "January", "February")
        trade_date: Date when the trade was opened (e.g., "12/22/2025", "1/6/2026")
        trade_time: Time when the trade was opened (e.g., "10:30 AM", "9:45 AM")
        sold_strike: Sold strike price with type prefix:
                     - "P 6855" for sold Put at $6855
                     - "C 6960" for sold Call at $6960
        delta: The delta value to record (numeric, e.g., 0.15, -0.08)
        delta_time: Time when the delta was obtained in EST (e.g., "10:45 AM", "2:30 PM")
                    Must fall within one of the accepted time windows:
                    - 9:30 AM - 11:00 AM
                    - 11:00 AM - 12:00 PM
                    - 1:00 PM - 2:00 PM
                    - 2:30 PM - 3:30 PM
    
    Returns:
        JSON string with operation result
    
    Column Mapping:
        - Column A: Date when the trade was opened (search column)
        - Column C: Time when the trade was opened (search column)
        - Column T: Delta for Sold Calls (9:30 AM - 11:00 AM)
        - Column U: Delta for Sold Puts (9:30 AM - 11:00 AM)
        - Column V: Delta for Sold Calls (11:00 AM - 12:00 PM)
        - Column W: Delta for Sold Puts (11:00 AM - 12:00 PM)
        - Column X: Delta for Sold Calls (1:00 PM - 2:00 PM)
        - Column Y: Delta for Sold Puts (1:00 PM - 2:00 PM)
        - Column Z: Delta for Sold Calls (2:30 PM - 3:30 PM)
        - Column AA: Delta for Sold Puts (2:30 PM - 3:30 PM)
    
    Example:
        excel_update_trade_with_delta(
            url="https://contoso.sharepoint.com/Shared%20Documents",
            file_name="2026 Trade Tracker.xlsx",
            sheet_name="January",
            trade_date="1/6/2026",
            trade_time="10:30 AM",
            sold_strike="P 6855",
            delta=0.12,
            delta_time="10:45 AM"
        )
        # This would update Column U (Sold Put delta for 9:30-11:00 window)
    """
    try:
        # Call the core implementation
        result = await update_trade_with_delta_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            trade_date=trade_date,
            trade_time=trade_time,
            sold_strike=sold_strike,
            delta=delta,
            delta_time=delta_time
        )
        
        return json.dumps(result, indent=2)
                
    except httpx.HTTPError as e:
        logger.error(f"HTTP error in updateTradeWithDelta: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error in updateTradeWithDelta: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


@mcp.tool(name="excel.closeTrade")
async def excel_close_trade(
    url: str,
    file_name: str,
    sheet_name: str,
    trade_date: str,
    trade_time: str,
    close_date: str,
    close_time: str
) -> str:
    """
    Close a trade by updating the close date and close time in an Excel workbook.
    
    This tool finds a trade row by matching the trade date (Column A) and trade time (Column C),
    then updates the close date (Column D) and close time (Column E) with the provided values.
    
    Args:
        url: SharePoint/OneDrive URL to the document library (e.g., https://contoso.sharepoint.com/Shared%20Documents)
        file_name: Excel workbook name with .xlsx extension (e.g., "2026 Trade Tracker.xlsx")
        sheet_name: Worksheet name (e.g., "January", "February")
        trade_date: Date when the trade was opened (e.g., "12/22/2025", "1/6/2026")
        trade_time: Time when the trade was opened (e.g., "10:30 AM", "9:45 AM")
        close_date: Date when the trade was closed (e.g., "12/22/2025", "1/6/2026")
        close_time: Time when the trade was closed (e.g., "4:00 PM", "2:30 PM")
    
    Returns:
        JSON string with operation result
    
    Column Mapping:
        - Column A: Date when the trade was opened (search column)
        - Column C: Time when the trade was opened (search column)
        - Column D: Date when the trade was closed (update column)
        - Column E: Time when the trade was closed (update column)
    
    Example:
        excel_close_trade(
            url="https://contoso.sharepoint.com/Shared%20Documents",
            file_name="2026 Trade Tracker.xlsx",
            sheet_name="January",
            trade_date="1/6/2026",
            trade_time="10:30 AM",
            close_date="1/6/2026",
            close_time="4:00 PM"
        )
        # This would update Column D with "1/6/2026" and Column E with "4:00 PM"
    """
    try:
        # Call the core implementation
        result = await close_trade_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            trade_date=trade_date,
            trade_time=trade_time,
            close_date=close_date,
            close_time=close_time
        )
        
        return json.dumps(result, indent=2)
                
    except httpx.HTTPError as e:
        logger.error(f"HTTP error in closeTrade: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error in closeTrade: {e}")
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
    url: str,
    file_name: str,
    sheet_name: str,
    trades: str
) -> str:
    """
    Log multiple trades to an Excel workbook.
    
    This tool automatically finds the last row with a valid date in column A and
    appends trade data starting from the next row.
    
    Note: This tool does NOT set close_date or close_time. Use the excel.closeTrade 
    tool to update those fields when a trade is closed.
    
    Args:
        url: SharePoint/OneDrive URL to the document library (e.g., https://contoso.sharepoint.com/Shared%20Documents)
        file_name: Excel workbook name with .xlsx extension (e.g., "2026 Trade Tracker.xlsx")
        sheet_name: Worksheet name (e.g., "January", "February")
        trades: JSON array of trade objects. Each object can have:
                - open_date: Date when trade was opened (e.g., "01/06/2026")
                - open_time: Time when trade was opened (e.g., "10:30 AM")
                - strategy: Strategy name (e.g., "VPCS", "VCCS", "IC", "Iron Condor")
                - credit: Credit received when opening (number, e.g., 0.25)
                - debit: Debit paid if closed before expiration (number, e.g., 0.10)
                - contracts: Number of contracts (integer, e.g., 25)
                - open_fees: Total fees paid when trade was opened (number, e.g., 88.78)
                - close_fees: Total fees paid if closed before expiration (number, e.g., 88.29)
                - sold_call_strike: Strike price for sold calls (number, e.g., 6960)
                - sold_put_strike: Strike price for sold puts (number, e.g., 6890)
                - width: Width in USD between sold and bought strikes (number, e.g., 15)
                
                Example: '[{"open_date": "01/06/2026", "open_time": "10:30 AM", "strategy": "VPCS", 
                           "credit": 0.30, "contracts": 25, "open_fees": 88.78, 
                           "sold_put_strike": 6890, "width": 15}]'
    
    Returns:
        JSON string with operation result including count of trades logged.
    
    Column Mapping:
        - Column A: Date when the trade was opened
        - Column C: Time when the trade was opened
        - Column G: Strategy
        - Column H: Credit Received
        - Column I: Debit Paid (only when trade is closed before expiration)
        - Column J: Number of Contracts
        - Column L: Total fees paid when the trade was opened
        - Column M: Total fees paid if the trade was closed before expiration
        - Column O: ATR (Average True Range)
        - Column P: Strike price for Sold Calls
        - Column Q: Strike price for Sold Puts
        - Column S: Width in USD between sold and bought strikes
   
    """
    try:
        # Parse the trades JSON string into a list
        try:
            trades_list = json.loads(trades)
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

        # Delegate to core implementation
        result = await log_trades_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            trades=trades_list
        )

        return json.dumps(result, indent=2)

    except Exception as e:
        logger.error(f"Error logging trades: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)


# =============================================================================
# REST API Endpoints (OpenAPI compatible)
# =============================================================================
# These endpoints provide a simpler REST interface alongside MCP.
# Faster than MCP protocol (single HTTP call vs session init + notification + tool call)
# OpenAPI docs available at /api/v1/openapi.json


class UpdateRangeRequest(BaseModel):
    """Request model for updateRange endpoint."""
    url: str = Field(..., description="SharePoint/OneDrive URL to the document library")
    file_name: str = Field(..., description="Excel file name with .xlsx extension")
    sheet_name: str = Field(..., description="Worksheet name")
    address: str = Field(..., description="Cell range address (e.g., 'A1:C3')")
    values: str = Field(..., description="JSON 2D array of values")


class LogTradesRequest(BaseModel):
    """Request model for logTrades endpoint."""
    url: str = Field(..., description="SharePoint/OneDrive URL to the document library")
    file_name: str = Field(..., description="Excel file name with .xlsx extension")
    sheet_name: str = Field(..., description="Worksheet name")
    trades: str = Field(..., description="JSON array of trade objects")


class UpdateDeltaRequest(BaseModel):
    """Request model for updateTradeWithDelta endpoint."""
    url: str = Field(..., description="SharePoint/OneDrive URL to the document library")
    file_name: str = Field(..., description="Excel file name with .xlsx extension")
    sheet_name: str = Field(..., description="Worksheet name")
    trade_date: str = Field(..., description="Date when the trade was opened")
    trade_time: str = Field(..., description="Time when the trade was opened")
    sold_strike: str = Field(..., description="Sold strike price with type prefix (e.g., 'P 6855')")
    delta: float = Field(..., description="The delta value to record")
    delta_time: str = Field(..., description="Time when the delta was obtained in EST")


class CloseTradeRequest(BaseModel):
    """Request model for closeTrade endpoint."""
    url: str = Field(..., description="SharePoint/OneDrive URL to the document library")
    file_name: str = Field(..., description="Excel file name with .xlsx extension")
    sheet_name: str = Field(..., description="Worksheet name")
    trade_date: str = Field(..., description="Date when the trade was opened")
    trade_time: str = Field(..., description="Time when the trade was opened")
    close_date: str = Field(..., description="Date when the trade was closed")
    close_time: str = Field(..., description="Time when the trade was closed")


@mcp.custom_route("/api/v1/updateRange", methods=["POST"])
async def api_update_range(request: Request) -> JSONResponse:
    """
    REST API: Update a range of cells in an Excel worksheet.
    
    Equivalent to the excel.updateRange MCP tool.
    """
    logger.info("REST API: updateRange")
    try:
        body = await request.json()
        
        # Validate required fields
        required_fields = ["url", "file_name", "sheet_name", "address", "values"]
        for field in required_fields:
            if field not in body:
                return JSONResponse(
                    {"status": "error", "message": f"Missing required field: {field}"},
                    status_code=400
                )
        
        # Parse values if it's a string
        values_data = body["values"]
        if isinstance(values_data, str):
            try:
                values_list = json.loads(values_data)
            except json.JSONDecodeError as e:
                return JSONResponse(
                    {"status": "error", "message": f"Invalid JSON in values: {str(e)}"},
                    status_code=400
                )
        else:
            values_list = values_data
        
        # Call the core implementation directly (not the MCP tool)
        result = await update_range_impl(
            url=body["url"],
            file_name=body["file_name"],
            sheet_name=body["sheet_name"],
            address=body["address"],
            values=values_list
        )
        
        status_code = 200 if result.get("status") == "success" else 400
        return JSONResponse(result, status_code=status_code)
        
    except json.JSONDecodeError as e:
        return JSONResponse(
            {"status": "error", "message": f"Invalid JSON in request body: {str(e)}"},
            status_code=400
        )
    except Exception as e:
        logger.error(f"REST API error (updateRange): {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@mcp.custom_route("/api/v1/logTrades", methods=["POST"])
async def api_log_trades(request: Request) -> JSONResponse:
    """
    REST API: Log multiple trades to an Excel workbook.
    
    Equivalent to the excel.logTrades MCP tool.
    """
    logger.info("REST API: logTrades")
    try:
        body = await request.json()
        
        # Validate required fields
        required_fields = ["url", "file_name", "sheet_name", "trades"]
        for field in required_fields:
            if field not in body:
                return JSONResponse(
                    {"status": "error", "message": f"Missing required field: {field}"},
                    status_code=400
                )
        
        # Parse trades if it's a string
        trades_data = body["trades"]
        if isinstance(trades_data, str):
            try:
                trades_list = json.loads(trades_data)
            except json.JSONDecodeError as e:
                return JSONResponse(
                    {"status": "error", "message": f"Invalid JSON in trades: {str(e)}"},
                    status_code=400
                )
        else:
            trades_list = trades_data
        
        # Call the core implementation directly (not the MCP tool)
        result = await log_trades_impl(
            url=body["url"],
            file_name=body["file_name"],
            sheet_name=body["sheet_name"],
            trades=trades_list
        )
        
        status_code = 200 if result.get("status") in ["success", "partial_success"] else 400
        return JSONResponse(result, status_code=status_code)
        
    except json.JSONDecodeError as e:
        return JSONResponse(
            {"status": "error", "message": f"Invalid JSON in request body: {str(e)}"},
            status_code=400
        )
    except Exception as e:
        logger.error(f"REST API error (logTrades): {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@mcp.custom_route("/api/v1/updateTradeWithDelta", methods=["POST"])
async def api_update_trade_with_delta(request: Request) -> JSONResponse:
    """
    REST API: Update delta value for a trade in an Excel workbook.
    
    Equivalent to the excel.updateTradeWithDelta MCP tool.
    
    Accepts alternative field names:
    - trade_date or open_date or executed_date
    - trade_time or open_time or lookup_value
    - sold_strike or strike_type (with strike_value)
    - delta or delta_value
    """
    logger.info("REST API: updateTradeWithDelta")
    try:
        body = await request.json()
        logger.info(f"REST API updateTradeWithDelta received body: {json.dumps(body)}")
        
        # Handle alternative field names
        # trade_date: can be "trade_date", "open_date", or "executed_date"
        trade_date = body.get("trade_date") or body.get("open_date") or body.get("executed_date") or ""
        
        # trade_time: can be "trade_time", "open_time", or "lookup_value"
        trade_time = body.get("trade_time") or body.get("open_time") or body.get("lookup_value") or ""
        
        # sold_strike: can be "sold_strike" or constructed from "strike_type" + "strike_value"
        sold_strike = body.get("sold_strike")
        if not sold_strike:
            strike_type = body.get("strike_type", "")
            strike_value = body.get("strike_value", "")
            if strike_type and strike_value:
                # Convert "sold_call" -> "C", "sold_put" -> "P"
                type_prefix = "C" if "call" in strike_type.lower() else "P" if "put" in strike_type.lower() else ""
                sold_strike = f"{type_prefix} {strike_value}" if type_prefix else ""
        
        # delta: can be "delta" or "delta_value"
        delta = body.get("delta") or body.get("delta_value")
        
        # delta_time is required
        delta_time = body.get("delta_time", "")
        
        # Validate required fields after mapping
        required_mappings = {
            "url": body.get("url"),
            "file_name": body.get("file_name"),
            "sheet_name": body.get("sheet_name"),
            "trade_date": trade_date,
            "trade_time": trade_time,
            "sold_strike": sold_strike,
            "delta": delta,
            "delta_time": delta_time
        }
        
        missing_fields = [k for k, v in required_mappings.items() if not v]
        if missing_fields:
            return JSONResponse(
                {"status": "error", "message": f"Missing required field(s): {', '.join(missing_fields)}. "
                 f"For trade_date, you can use: trade_date, open_date, or executed_date. "
                 f"For trade_time, you can use: trade_time, open_time, or lookup_value. "
                 f"For sold_strike, you can use: sold_strike or (strike_type + strike_value). "
                 f"For delta, you can use: delta or delta_value."},
                status_code=400
            )
        
        # Call the core implementation directly (not the MCP tool)
        result = await update_trade_with_delta_impl(
            url=body["url"],
            file_name=body["file_name"],
            sheet_name=body["sheet_name"],
            trade_date=trade_date,
            trade_time=trade_time,
            sold_strike=sold_strike,
            delta=float(delta),
            delta_time=delta_time
        )
        
        status_code = 200 if result.get("status") == "success" else 400
        return JSONResponse(result, status_code=status_code)
        
    except json.JSONDecodeError as e:
        return JSONResponse(
            {"status": "error", "message": f"Invalid JSON in request body: {str(e)}"},
            status_code=400
        )
    except Exception as e:
        logger.error(f"REST API error (updateTradeWithDelta): {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@mcp.custom_route("/api/v1/closeTrade", methods=["POST"])
async def api_close_trade(request: Request) -> JSONResponse:
    """
    REST API: Close a trade by updating close date and time.
    
    Equivalent to the excel.closeTrade MCP tool.
    
    Accepts alternative field names:
    - trade_date or open_date or executed_date
    - trade_time or open_time or executed_time
    - close_date or closed_date
    - close_time or closed_time
    """
    logger.info("REST API: closeTrade")
    try:
        body = await request.json()
        logger.info(f"REST API closeTrade received body: {json.dumps(body)}")
        
        # Handle alternative field names
        trade_date = body.get("trade_date") or body.get("open_date") or body.get("executed_date") or ""
        trade_time = body.get("trade_time") or body.get("open_time") or body.get("executed_time") or ""
        close_date = body.get("close_date") or body.get("closed_date") or ""
        close_time = body.get("close_time") or body.get("closed_time") or ""
        
        # Validate required fields after mapping
        required_mappings = {
            "url": body.get("url"),
            "file_name": body.get("file_name"),
            "sheet_name": body.get("sheet_name"),
            "trade_date": trade_date,
            "trade_time": trade_time,
            "close_date": close_date,
            "close_time": close_time
        }
        
        missing_fields = [k for k, v in required_mappings.items() if not v]
        if missing_fields:
            return JSONResponse(
                {"status": "error", "message": f"Missing required field(s): {', '.join(missing_fields)}"},
                status_code=400
            )
        
        # Call the core implementation directly (not the MCP tool)
        result = await close_trade_impl(
            url=body["url"],
            file_name=body["file_name"],
            sheet_name=body["sheet_name"],
            trade_date=trade_date,
            trade_time=trade_time,
            close_date=close_date,
            close_time=close_time
        )
        
        status_code = 200 if result.get("status") == "success" else 400
        return JSONResponse(result, status_code=status_code)
        
    except json.JSONDecodeError as e:
        return JSONResponse(
            {"status": "error", "message": f"Invalid JSON in request body: {str(e)}"},
            status_code=400
        )
    except Exception as e:
        logger.error(f"REST API error (closeTrade): {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@mcp.custom_route("/api/v1/openapi.json", methods=["GET"])
async def api_openapi_spec(request: Request) -> JSONResponse:
    """Return OpenAPI 3.0 specification for the REST API."""
    spec = {
        "openapi": "3.0.0",
        "info": {
            "title": "MCP Excel Service REST API",
            "description": "REST API for Excel operations via Microsoft Graph. Provides the same functionality as the MCP tools but via standard REST endpoints.",
            "version": "1.0.0"
        },
        "servers": [
            {"url": "/api/v1", "description": "REST API v1"}
        ],
        "paths": {
            "/updateRange": {
                "post": {
                    "summary": "Update Range",
                    "description": "Update a range of cells in an Excel worksheet",
                    "operationId": "updateRange",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/UpdateRangeRequest"}
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Range updated successfully",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/OperationResponse"}
                                }
                            }
                        },
                        "400": {"description": "Bad request or operation failed"},
                        "401": {"description": "Unauthorized - missing or invalid bearer token"},
                        "500": {"description": "Internal server error"}
                    }
                }
            },
            "/logTrades": {
                "post": {
                    "summary": "Log Trades",
                    "description": "Log multiple trades to an Excel workbook",
                    "operationId": "logTrades",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/LogTradesRequest"}
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Trades logged successfully",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/LogTradesResponse"}
                                }
                            }
                        },
                        "400": {"description": "Bad request or operation failed"},
                        "401": {"description": "Unauthorized - missing or invalid bearer token"},
                        "500": {"description": "Internal server error"}
                    }
                }
            },
            "/updateTradeWithDelta": {
                "post": {
                    "summary": "Update Trade with Delta",
                    "description": "Update delta value for a trade in an Excel workbook",
                    "operationId": "updateTradeWithDelta",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/UpdateDeltaRequest"}
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Delta updated successfully",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/OperationResponse"}
                                }
                            }
                        },
                        "400": {"description": "Bad request or operation failed"},
                        "401": {"description": "Unauthorized - missing or invalid bearer token"},
                        "500": {"description": "Internal server error"}
                    }
                }
            },
            "/closeTrade": {
                "post": {
                    "summary": "Close Trade",
                    "description": "Close a trade by updating close date and time in an Excel workbook",
                    "operationId": "closeTrade",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/CloseTradeRequest"}
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Trade closed successfully",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/OperationResponse"}
                                }
                            }
                        },
                        "400": {"description": "Bad request or operation failed"},
                        "401": {"description": "Unauthorized - missing or invalid bearer token"},
                        "500": {"description": "Internal server error"}
                    }
                }
            }
        },
        "components": {
            "schemas": {
                "UpdateRangeRequest": {
                    "type": "object",
                    "required": ["url", "file_name", "sheet_name", "address", "values"],
                    "properties": {
                        "url": {"type": "string", "description": "SharePoint/OneDrive URL to the document library", "example": "https://contoso.sharepoint.com/Shared%20Documents"},
                        "file_name": {"type": "string", "description": "Excel file name with .xlsx extension", "example": "Budget.xlsx"},
                        "sheet_name": {"type": "string", "description": "Worksheet name", "example": "Sheet1"},
                        "address": {"type": "string", "description": "Cell range address", "example": "A1:C3"},
                        "values": {"type": "string", "description": "JSON 2D array of values", "example": "[[\"a\", \"b\"], [\"c\", \"d\"]]"}
                    }
                },
                "LogTradesRequest": {
                    "type": "object",
                    "required": ["url", "file_name", "sheet_name", "trades"],
                    "properties": {
                        "url": {"type": "string", "description": "SharePoint/OneDrive URL", "example": "https://contoso.sharepoint.com/Shared%20Documents"},
                        "file_name": {"type": "string", "description": "Excel file name", "example": "2026 Trade Tracker.xlsx"},
                        "sheet_name": {"type": "string", "description": "Worksheet name", "example": "January"},
                        "trades": {"type": "string", "description": "JSON array of trade objects", "example": "[{\"open_date\": \"01/06/2026\", \"strategy\": \"VPCS\", \"credit\": 0.30}]"}
                    }
                },
                "UpdateDeltaRequest": {
                    "type": "object",
                    "required": ["url", "file_name", "sheet_name", "trade_date", "trade_time", "sold_strike", "delta", "delta_time"],
                    "properties": {
                        "url": {"type": "string", "description": "SharePoint/OneDrive URL"},
                        "file_name": {"type": "string", "description": "Excel file name"},
                        "sheet_name": {"type": "string", "description": "Worksheet name"},
                        "trade_date": {"type": "string", "description": "Date when trade was opened", "example": "1/6/2026"},
                        "trade_time": {"type": "string", "description": "Time when trade was opened", "example": "10:30 AM"},
                        "sold_strike": {"type": "string", "description": "Sold strike with type prefix", "example": "P 6855"},
                        "delta": {"type": "number", "description": "Delta value to record", "example": 0.12},
                        "delta_time": {"type": "string", "description": "Time when delta was obtained", "example": "10:45 AM"}
                    }
                },
                "CloseTradeRequest": {
                    "type": "object",
                    "required": ["url", "file_name", "sheet_name", "trade_date", "trade_time", "close_date", "close_time"],
                    "properties": {
                        "url": {"type": "string", "description": "SharePoint/OneDrive URL"},
                        "file_name": {"type": "string", "description": "Excel file name"},
                        "sheet_name": {"type": "string", "description": "Worksheet name"},
                        "trade_date": {"type": "string", "description": "Date when trade was opened", "example": "1/6/2026"},
                        "trade_time": {"type": "string", "description": "Time when trade was opened", "example": "10:30 AM"},
                        "close_date": {"type": "string", "description": "Date when trade was closed", "example": "1/6/2026"},
                        "close_time": {"type": "string", "description": "Time when trade was closed", "example": "4:00 PM"}
                    }
                },
                "OperationResponse": {
                    "type": "object",
                    "properties": {
                        "status": {"type": "string", "enum": ["success", "error", "partial_success"]},
                        "message": {"type": "string"},
                        "file_name": {"type": "string"},
                        "sheet_name": {"type": "string"}
                    }
                },
                "LogTradesResponse": {
                    "type": "object",
                    "properties": {
                        "status": {"type": "string"},
                        "message": {"type": "string"},
                        "trades_logged": {"type": "integer"},
                        "file_name": {"type": "string"},
                        "sheet_name": {"type": "string"},
                        "start_row": {"type": "integer"},
                        "results": {"type": "array", "items": {"type": "object"}}
                    }
                }
            },
            "securitySchemes": {
                "bearerAuth": {
                    "type": "http",
                    "scheme": "bearer",
                    "bearerFormat": "JWT",
                    "description": "Microsoft Entra ID bearer token. Obtain via client credentials flow or managed identity."
                }
            }
        },
        "security": [
            {"bearerAuth": []}
        ]
    }
    return JSONResponse(spec)


# =============================================================================
# Health Check Endpoint
# =============================================================================

@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Health check endpoint for Container Apps."""
    auth_status = "enabled" if ENTRA_AUTH_ENABLED else "disabled"
    return JSONResponse({
        "status": "healthy",
        "service": "MCP Excel Service",
        "version": "2.0.0",
        "transport": "streamable-http",
        "mcp_tools": [
            "excel.updateRange",
            "excel.logTrades",
            "excel.updateTradeWithDelta",
            "excel.closeTrade"
        ],
        "rest_api": {
            "base_path": "/api/v1",
            "endpoints": [
                "POST /api/v1/updateRange",
                "POST /api/v1/logTrades",
                "POST /api/v1/updateTradeWithDelta",
                "POST /api/v1/closeTrade",
                "GET /api/v1/openapi.json"
            ]
        },
        "authentication": auth_status,
        "tenant_id": ENTRA_TENANT_ID[:8] + "..." if ENTRA_TENANT_ID else "not configured",
        "client_id": ENTRA_CLIENT_ID[:8] + "..." if ENTRA_CLIENT_ID else "not configured"
    })


# =============================================================================
# Server Entry Point
# =============================================================================

if __name__ == "__main__":
    import uvicorn
    from starlette.applications import Starlette
    from starlette.routing import Mount
    
    port = int(os.getenv("PORT", "3000"))
    host = os.getenv("HOST", "0.0.0.0")
    
    # Configure and log authentication settings
    configure_auth_middleware()
    
    logger.info(f"Starting MCP Excel Service on {host}:{port}")
    logger.info(f"MCP endpoint: http://{host}:{port}/mcp")
    logger.info(f"REST API: http://{host}:{port}/api/v1/*")
    logger.info(f"OpenAPI spec: http://{host}:{port}/api/v1/openapi.json")
    logger.info(f"Health endpoint: http://{host}:{port}/health")
    
    # Check for Entra auth configuration
    validator = create_validator_from_env()
    if validator:
        logger.info("Microsoft Entra ID authentication ENABLED")
        logger.info("  - Health endpoint (/health) is NOT protected")
        logger.info("  - All other endpoints require bearer token")

        # Get the HTTP app from FastMCP
        http_app = mcp.http_app()
        
        # Create a wrapper Starlette app that uses FastMCP's lifespan
        # This ensures the task group is properly initialized
        wrapper_app = Starlette(
            routes=[Mount("/", app=http_app)],
            lifespan=http_app.lifespan  # Critical: pass the lifespan!
        )
        
        # Wrap with auth middleware
        app_with_auth = EntraAuthMiddleware(
            wrapper_app,
            validator=validator,
            excluded_paths=["/health"]
        )

        # Run with uvicorn
        uvicorn.run(app_with_auth, host=host, port=port)
    else:
        logger.info("Microsoft Entra ID authentication DISABLED (no credentials configured)")
        # Run normally without auth
        mcp.run(transport="http", host=host, port=port)
