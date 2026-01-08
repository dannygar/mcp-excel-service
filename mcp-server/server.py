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
"""

import os
import json
import logging
import pathlib
from datetime import datetime

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
from core_operations import update_trade_with_delta_impl, close_trade_impl

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
            
            # Get open date/time
            open_date = trade.get("open_date", trade.get("date", ""))
            open_time = trade.get("open_time", trade.get("time", ""))
            debit = trade.get("debit", "")
            
            # Build values dictionary matching COLUMN_MAP
            # Note: close_date/close_time are NOT set here - use excel.closeTrade tool
            values_dict = {
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
            }
            
            # Log the values for debugging
            logger.info(f"Trade {i+1} values: C={values_dict['open_date']}, E={values_dict['open_time']}, I={values_dict['strategy']}, J={values_dict['credit']}")
            
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
# Health Check Endpoint
# =============================================================================

@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Health check endpoint for Container Apps."""
    auth_status = "enabled" if ENTRA_AUTH_ENABLED else "disabled"
    return JSONResponse({
        "status": "healthy",
        "service": "mcp-excel-server",
        "authentication": auth_status,
        "tenant_id": ENTRA_TENANT_ID[:8] + "..." if ENTRA_TENANT_ID else "not configured",
        "client_id": ENTRA_CLIENT_ID[:8] + "..." if ENTRA_CLIENT_ID else "not configured"
    })


# =============================================================================
# Server Entry Point
# =============================================================================

if __name__ == "__main__":
    from starlette.middleware import Middleware
    
    port = int(os.getenv("PORT", "3000"))
    host = os.getenv("HOST", "0.0.0.0")
    
    # Configure and log authentication settings
    configure_auth_middleware()
    
    logger.info(f"Starting MCP Excel Service on {host}:{port}")
    logger.info(f"MCP endpoint: http://{host}:{port}/mcp")
    logger.info(f"Health endpoint: http://{host}:{port}/health")
    
    # Build middleware list based on authentication configuration
    middleware = []
    if ENTRA_AUTH_ENABLED:
        logger.info("Adding Entra ID authentication middleware...")
        middleware.append(Middleware(EntraAuthMiddleware))
    
    # Create ASGI app with middleware and run it
    # Using http_app() allows us to properly configure middleware
    app = mcp.http_app(middleware=middleware if middleware else None)
    
    # Run with uvicorn directly (fastmcp wraps this)
    import uvicorn
    uvicorn.run(app, host=host, port=port)
