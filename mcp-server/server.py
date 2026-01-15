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
    
    This tool finds a trade row by matching the trade date (Column C) and trade time (Column E),
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
        - Column C: Date when the trade was opened (search column)
        - Column E: Time when the trade was opened (search column)
        - Column U: Delta for Sold Calls (9:30 AM - 11:00 AM)
        - Column V: Delta for Sold Puts (9:30 AM - 11:00 AM)
        - Column W: Delta for Sold Calls (11:00 AM - 12:00 PM)
        - Column X: Delta for Sold Puts (11:00 AM - 12:00 PM)
        - Column Y: Delta for Sold Calls (1:00 PM - 2:00 PM)
        - Column Z: Delta for Sold Puts (1:00 PM - 2:00 PM)
        - Column AA: Delta for Sold Calls (2:30 PM - 3:30 PM)
        - Column AB: Delta for Sold Puts (2:30 PM - 3:30 PM)
    
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
        # This would update Column V (Sold Put delta for 9:30-11:00 window)
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
    
    This tool finds a trade row by matching the trade date (Column C) and trade time (Column E),
    then updates the close date (Column F) and close time (Column G) with the provided values.
    
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
        - Column C: Date when the trade was opened (search column)
        - Column E: Time when the trade was opened (search column)
        - Column F: Date when the trade was closed (update column)
        - Column G: Time when the trade was closed (update column)
    
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
        # This would update Column F with "1/6/2026" and Column G with "4:00 PM"
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
    
    This tool automatically finds the last row with a valid date in column C and
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
        - Column C: Date when the trade was opened
        - Column E: Time when the trade was opened
        - Column I: Strategy
        - Column J: Credit Received
        - Column K: Debit Paid (only when trade is closed before expiration)
        - Column L: Number of Contracts
        - Column N: Total fees paid when the trade was opened
        - Column O: Total fees paid if the trade was closed before expiration
        - Column Q: Strike price for Sold Calls
        - Column R: Strike price for Sold Puts
        - Column T: Width in USD between sold and bought strikes
   
    """
    # Column configuration for trade tracker
    # C=open_date, E=open_time, I=strategy, J=credit, K=debit, L=contracts, 
    # N=open_fees, O=close_fees, Q=sold_call, R=sold_put, T=width
    # Note: Close date/time (F, G) are NOT set here - use excel.closeTrade tool instead
    COLUMN_MAP = {
        "C": "open_date",
        "E": "open_time",
        "I": "strategy",
        "J": "credit",
        "K": "debit",
        "L": "contracts",
        "N": "open_fees",
        "O": "close_fees",
        "Q": "sold_call_strike",
        "R": "sold_put_strike",
        "T": "width",
    }
    
    # Log incoming parameters for debugging
    logger.info(f"excel.logTrades called with sheet_name='{sheet_name}'")
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
            # Handle alternative field names from different clients
            open_date_str = trade.get("open_date") or trade.get("date") or trade.get("executed_date") or ""
            open_time_str = trade.get("open_time") or trade.get("time") or trade.get("executed_time") or ""
            
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
        
        # Find the last non-empty date in column C to determine where to insert new rows
        logger.info(f"Finding last date in column C of sheet '{sheet_name}'")
        
        # Resolve URL to get drive_id, item_id, and site_id
        resolved = await resolve_excel_file_ids(url, file_name)
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
            last_date_row = None
            
            for i in range(len(column_values) - 1, -1, -1):
                cell_value = column_values[i][0] if column_values[i] else None
                if cell_value is not None and cell_value != "":
                    # Check if it's a valid date (either Excel serial number or date string)
                    try:
                        # If it's a number (Excel serial date), convert to date string
                        if isinstance(cell_value, (int, float)) and 1 <= cell_value <= 2958465:
                            dt = excel_serial_to_date(cell_value)
                            last_date_display = dt.strftime("%m/%d/%Y")
                            last_date_row = i + 1
                            logger.info(f"Found last date at row {last_date_row}: Excel serial {cell_value} → {last_date_display}")
                            break
                        # If it's a string, try to parse it as a date
                        elif isinstance(cell_value, str):
                            parsed = parse_date_string(cell_value)
                            if parsed:
                                last_date_row = i + 1
                                logger.info(f"Found last date at row {last_date_row}: {cell_value}")
                                break
                    except (ValueError, TypeError):
                        continue
            
            # Fallback: If no valid date found, look for "Date of Purchase" header cell
            if last_date_row is None:
                logger.info("No valid dates found in column C, searching for 'Date of Purchase' header...")
                for i, row_value in enumerate(column_values):
                    cell_value = row_value[0] if row_value else None
                    if cell_value is not None and isinstance(cell_value, str):
                        # Check for "Date of Purchase" (case-insensitive, partial match)
                        if "date" in cell_value.lower() and "purchase" in cell_value.lower():
                            last_date_row = i + 1  # Use the header row as the reference
                            logger.info(f"Found 'Date of Purchase' header at row {last_date_row}, will start writing at row {last_date_row + 1}")
                            break
            
            if last_date_row is None:
                return json.dumps({
                    "status": "error",
                    "message": f"Could not find any valid date or 'Date of Purchase' header in column C of sheet '{sheet_name}'",
                }, indent=2)
        
        logger.info(f"Logging {len(trades_list)} trades to {file_name}, sheet '{sheet_name}'")
        logger.info(f"Last date row: {last_date_row}, will start writing at row {last_date_row + 1}")
        
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
            
            # Handle alternative field names from different clients
            # open_date: can be "open_date", "date", or "executed_date"
            open_date = trade.get("open_date") or trade.get("date") or trade.get("executed_date") or ""
            
            # open_time: can be "open_time", "time", or "executed_time"
            open_time = trade.get("open_time") or trade.get("time") or trade.get("executed_time") or ""
            
            # credit: can be "credit" or "credit_received" (convert to per-contract if needed)
            credit = trade.get("credit")
            if credit is None or credit == "":
                credit_received = trade.get("credit_received")
                contracts = trade.get("contracts", 1)
                if credit_received and contracts:
                    # credit_received is total, convert to per-contract (divided by 100 for options)
                    credit = credit_received / (contracts * 100) if contracts > 0 else ""
                else:
                    credit = ""
            
            # debit: can be "debit" or "debit_paid" (convert to per-contract if needed)
            debit = trade.get("debit")
            if debit is None or debit == "":
                debit_paid = trade.get("debit_paid")
                contracts = trade.get("contracts", 1)
                if debit_paid and contracts:
                    debit = debit_paid / (contracts * 100) if contracts > 0 else ""
                else:
                    debit = ""
            
            # open_fees: can be "open_fees", "fees", or "total_fees"
            open_fees = trade.get("open_fees") or trade.get("fees") or trade.get("total_fees") or ""
            
            # sold_call_strike: can be "sold_call_strike" or extracted from "sold_strikes" array for calls
            sold_call_strike = trade.get("sold_call_strike")
            if (sold_call_strike is None or sold_call_strike == "") and trade.get("sold_strikes"):
                # Check if this is a call spread (strategy contains "Call" or only calls_contracts > 0)
                is_call_spread = (
                    "call" in raw_strategy.lower() or 
                    (trade.get("calls_contracts", 0) > 0 and trade.get("puts_contracts", 0) == 0)
                )
                if is_call_spread and trade.get("sold_strikes"):
                    sold_call_strike = trade["sold_strikes"][0] if trade["sold_strikes"] else ""
            
            # sold_put_strike: can be "sold_put_strike" or extracted from "sold_strikes" array for puts
            sold_put_strike = trade.get("sold_put_strike")
            if (sold_put_strike is None or sold_put_strike == "") and trade.get("sold_strikes"):
                # Check if this is a put spread (strategy contains "Put" or only puts_contracts > 0)
                is_put_spread = (
                    "put" in raw_strategy.lower() or 
                    (trade.get("puts_contracts", 0) > 0 and trade.get("calls_contracts", 0) == 0)
                )
                if is_put_spread and trade.get("sold_strikes"):
                    sold_put_strike = trade["sold_strikes"][0] if trade["sold_strikes"] else ""
            
            # width: can be "width" or calculated from sold_strikes and bought_strikes
            width = trade.get("width")
            if (width is None or width == "") and trade.get("sold_strikes") and trade.get("bought_strikes"):
                sold = trade["sold_strikes"][0] if trade["sold_strikes"] else None
                bought = trade["bought_strikes"][0] if trade["bought_strikes"] else None
                if sold is not None and bought is not None:
                    width = abs(bought - sold)
            
            # Build values dictionary matching COLUMN_MAP
            # Note: close_date/close_time are NOT set here - use excel.closeTrade tool
            values_dict = {
                "open_date": open_date,
                "open_time": open_time,
                "strategy": mapped_strategy,
                "credit": credit,
                "debit": debit,
                "contracts": trade.get("contracts", ""),
                "open_fees": open_fees,
                "close_fees": trade.get("close_fees", ""),
                "sold_call_strike": sold_call_strike if sold_call_strike else "",
                "sold_put_strike": sold_put_strike if sold_put_strike else "",
                "width": width if width else "",
            }
            
            # Log the extracted values for debugging
            logger.info(f"Trade {i+1} extracted values: open_date={open_date}, open_time={open_time}, "
                        f"strategy={mapped_strategy}, credit={credit}, contracts={trade.get('contracts')}, "
                        f"open_fees={open_fees}, sold_call={sold_call_strike}, sold_put={sold_put_strike}, width={width}")
            
            # Calculate target row: first trade goes to last_date_row + 1, second to last_date_row + 2, etc.
            target_row = last_date_row + 1 + i
            
            logger.info(f"Logging trade {i+1}/{len(trades_list)}: strategy={mapped_strategy}, credit={trade.get('credit')}, contracts={trade.get('contracts')} to row {target_row}")
            
            # Write each cell directly to the calculated row
            updated_cells = []
            cell_errors = []
            
            async with httpx.AsyncClient() as write_client:
                for col_letter, field_name in COLUMN_MAP.items():
                    value = values_dict.get(field_name, "")
                    # Skip empty values
                    if value == "" or value is None:
                        continue
                    
                    cell_address = f"{col_letter}{target_row}"
                    cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{cell_address}')"
                    
                    try:
                        response = await write_client.patch(
                            cell_url,
                            headers=headers,
                            json={"values": [[value]]},
                            timeout=30.0
                        )
                        
                        if response.status_code in [200, 201]:
                            updated_cells.append(cell_address)
                        else:
                            error_data = response.json() if response.content else {}
                            error_message = error_data.get("error", {}).get("message", response.text)
                            cell_errors.append(f"{cell_address}: {error_message}")
                    except Exception as cell_error:
                        cell_errors.append(f"{col_letter}{target_row}: {str(cell_error)}")
            
            if cell_errors:
                errors.append({
                    "trade_index": i + 1,
                    "error": f"Failed to update some cells: {', '.join(cell_errors)}",
                    "strategy": mapped_strategy,
                    "updated_cells": updated_cells,
                })
            else:
                results.append({
                    "trade_index": i + 1,
                    "row": target_row,
                    "open_date": open_date,
                    "open_time": open_time,
                    "strategy": mapped_strategy,
                    "credit": trade.get("credit", ""),
                    "debit": debit,
                    "contracts": trade.get("contracts", ""),
                    "open_fees": trade.get("open_fees", trade.get("fees", "")),
                    "close_fees": trade.get("close_fees", ""),
                    "sold_call_strike": trade.get("sold_call_strike", ""),
                    "sold_put_strike": trade.get("sold_put_strike", ""),
                    "width": trade.get("width", ""),
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
                "file_name": file_name,
                "sheet_name": sheet_name,
                "start_row": last_date_row + 1,
                "results": results,
                "errors": errors,
            }, indent=2)
        else:
            return json.dumps({
                "status": "success",
                "message": f"Successfully logged {len(results)} trades",
                "trades_logged": len(results),
                "file_name": file_name,
                "sheet_name": sheet_name,
                "start_row": last_date_row + 1,
                "results": results,
            }, indent=2)
            
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
