"""
Core Excel Operations Module

Provides the core implementation functions for Excel operations.
These functions are called by the MCP tools in server.py.
"""

import json
import logging
from datetime import datetime

import httpx

from graph_api import get_graph_headers, build_workbook_url
from excel_helpers import resolve_excel_file_ids, compare_values_for_search, parse_date_string, excel_serial_to_date
from config import map_strategy_name

logger = logging.getLogger("mcp-excel-server")


# =============================================================================
# Delta Column Mapping
# =============================================================================
# Maps (time_window, strike_type) to column letter
# Time windows: "9:30-11:00", "11:00-12:00", "1:00-2:00", "2:30-3:30"
# Strike types: "C" (Call), "P" (Put)
DELTA_COLUMN_MAPPING = {
    ("9:30-11:00", "C"): "U",   # Sold Calls 9:30am-11:00am
    ("9:30-11:00", "P"): "V",   # Sold Puts 9:30am-11:00am
    ("11:00-12:00", "C"): "W",  # Sold Calls 11:00am-12:00pm
    ("11:00-12:00", "P"): "X",  # Sold Puts 11:00am-12:00pm
    ("1:00-2:00", "C"): "Y",    # Sold Calls 1:00pm-2:00pm
    ("1:00-2:00", "P"): "Z",    # Sold Puts 1:00pm-2:00pm
    ("2:30-3:30", "C"): "AA",   # Sold Calls 2:30pm-3:30pm
    ("2:30-3:30", "P"): "AB",   # Sold Puts 2:30pm-3:30pm
}

# Time window boundaries (in minutes from midnight)
TIME_WINDOWS = [
    # (start_minutes, end_minutes, window_key)
    (9 * 60 + 30, 11 * 60, "9:30-11:00"),      # 9:30 AM - 11:00 AM
    (11 * 60, 12 * 60, "11:00-12:00"),          # 11:00 AM - 12:00 PM
    (13 * 60, 14 * 60, "1:00-2:00"),            # 1:00 PM - 2:00 PM
    (14 * 60 + 30, 15 * 60 + 30, "2:30-3:30"),  # 2:30 PM - 3:30 PM
]


def parse_time_to_minutes(time_str: str) -> int | None:
    """
    Parse a time string to minutes from midnight.
    Supports formats like "10:30 AM", "2:30 PM", "14:30", etc.
    
    Returns None if parsing fails.
    """
    time_formats = [
        "%I:%M %p",      # 10:30 AM
        "%I:%M%p",       # 10:30AM
        "%H:%M",         # 14:30
        "%H:%M:%S",      # 14:30:00
        "%I:%M:%S %p",   # 10:30:00 AM
    ]
    
    for fmt in time_formats:
        try:
            parsed = datetime.strptime(time_str.strip().upper(), fmt)
            return parsed.hour * 60 + parsed.minute
        except ValueError:
            continue
    
    return None


def get_time_window(delta_time: str) -> str | None:
    """
    Determine which time window a given delta_time falls into.
    
    Returns the window key (e.g., "9:30-11:00") or None if not in any window.
    """
    minutes = parse_time_to_minutes(delta_time)
    if minutes is None:
        return None
    
    for start, end, window_key in TIME_WINDOWS:
        if start <= minutes < end:
            return window_key
    
    return None


def parse_sold_strike(sold_strike: str) -> tuple[str, str] | None:
    """
    Parse the sold strike string to extract type and strike price.
    
    Args:
        sold_strike: String like "P 6855" or "C 6960"
    
    Returns:
        Tuple of (strike_type, strike_price) or None if invalid format.
        strike_type is "C" for Call or "P" for Put.
    """
    parts = sold_strike.strip().upper().split()
    if len(parts) != 2:
        return None
    
    strike_type = parts[0]
    if strike_type not in ("C", "P"):
        return None
    
    strike_price = parts[1]
    return (strike_type, strike_price)


# =============================================================================
# Core Implementation Functions (called by tools)
# =============================================================================

async def update_trade_with_delta_impl(
    url: str,
    file_name: str,
    sheet_name: str,
    trade_date: str,
    trade_time: str,
    sold_strike: str,
    delta: float,
    delta_time: str
) -> dict:
    """
    Core implementation for updating trade delta values.
    Returns a dict with the result (not a JSON string).
    
    Args:
        url: SharePoint/OneDrive URL to the document library
        file_name: Excel file name with .xlsx extension
        sheet_name: Worksheet name
        trade_date: Date when the trade was opened (Column C)
        trade_time: Time when the trade was opened (Column E)
        sold_strike: Sold strike with type prefix (e.g., "P 6855", "C 6960")
        delta: The delta value to record
        delta_time: Time when the delta was obtained (determines which column)
    
    Returns:
        Dictionary with operation result
    """
    # Parse and validate the sold strike
    strike_parsed = parse_sold_strike(sold_strike)
    if strike_parsed is None:
        return {
            "status": "error",
            "message": f"Invalid sold_strike format: '{sold_strike}'. Expected format: 'P 6855' (Put) or 'C 6960' (Call)",
        }
    strike_type, strike_price = strike_parsed
    
    # Determine the time window for the delta_time
    time_window = get_time_window(delta_time)
    if time_window is None:
        accepted_windows = [
            "9:30 AM - 11:00 AM",
            "11:00 AM - 12:00 PM",
            "1:00 PM - 2:00 PM",
            "2:30 PM - 3:30 PM",
        ]
        return {
            "status": "error",
            "message": f"Delta time '{delta_time}' does not fall into any accepted time window",
            "accepted_windows": accepted_windows,
            "hint": "Ensure the delta_time is in EST and falls within one of the accepted windows",
        }
    
    # Get the target column based on time window and strike type
    target_column = DELTA_COLUMN_MAPPING.get((time_window, strike_type))
    if target_column is None:
        return {
            "status": "error",
            "message": f"No column mapping found for time window '{time_window}' and strike type '{strike_type}'",
        }
    
    logger.info(f"Delta update: trade_date={trade_date}, trade_time={trade_time}, "
                f"strike={sold_strike}, delta={delta}, delta_time={delta_time} -> Column {target_column}")
    
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
                "message": f"Worksheet '{sheet_name}' is empty",
            }
        
        # Step 2: Get columns C (date) and E (time) to find the matching row
        # Read both columns in one request for efficiency
        search_range = f"C1:E{row_count}"
        search_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{search_range}')"
        logger.info(f"Searching for trade with date='{trade_date}' and time='{trade_time}'")
        
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
                "message": f"Failed to read search columns: {error_message}",
                "status_code": search_response.status_code,
            }
        
        search_data = search_response.json()
        row_values = search_data.get("values", [])
        
        # Step 3: Find the row with matching date (C) and time (E)
        # Note: C is index 0, D is index 1, E is index 2
        found_row = None
        for i, row in enumerate(row_values):
            date_value = row[0] if len(row) > 0 else None  # Column C
            time_value = row[2] if len(row) > 2 else None  # Column E
            
            # Compare date
            date_match = compare_values_for_search(date_value, trade_date)
            
            # Compare time (normalize both values for comparison)
            time_match = False
            if time_value is not None and trade_time:
                # Normalize time values for comparison
                time_value_str = str(time_value).strip().upper()
                trade_time_str = trade_time.strip().upper()
                # Direct string comparison or parse and compare
                if time_value_str == trade_time_str:
                    time_match = True
                else:
                    # Try parsing both times and comparing
                    time_value_minutes = parse_time_to_minutes(str(time_value))
                    trade_time_minutes = parse_time_to_minutes(trade_time)
                    if time_value_minutes is not None and trade_time_minutes is not None:
                        time_match = (time_value_minutes == trade_time_minutes)
            
            if date_match and time_match:
                found_row = i + 1  # Excel rows are 1-indexed
                logger.info(f"Found matching trade at row {found_row}: date='{date_value}', time='{time_value}'")
                break
        
        if found_row is None:
            # Provide diagnostic info
            sample_rows = []
            for i, row in enumerate(row_values[:10]):
                date_val = row[0] if len(row) > 0 else "empty"
                time_val = row[2] if len(row) > 2 else "empty"
                sample_rows.append(f"Row {i+1}: date='{date_val}', time='{time_val}'")
            
            return {
                "status": "error",
                "message": f"No trade found with date '{trade_date}' and time '{trade_time}' in sheet '{sheet_name}'",
                "searched_rows": row_count,
                "sample_rows": sample_rows,
                "hint": "Ensure the trade_date and trade_time match exactly what's in the spreadsheet",
            }
        
        # Step 4: Update the delta value in the target column
        cell_address = f"{target_column}{found_row}"
        cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{cell_address}')"
        
        body = {
            "values": [[delta]],
        }
        
        logger.info(f"Updating cell '{cell_address}' with delta value '{delta}'")
        
        update_response = await client.patch(
            cell_url,
            headers=headers,
            json=body,
            timeout=30.0,
        )
        
        if update_response.status_code != 200:
            error_data = update_response.json() if update_response.content else {}
            error_message = error_data.get("error", {}).get("message", update_response.text)
            return {
                "status": "error",
                "message": f"Failed to update cell {cell_address}: {error_message}",
                "status_code": update_response.status_code,
            }
        
        strike_type_name = "Put" if strike_type == "P" else "Call"
        return {
            "status": "success",
            "message": f"Successfully updated {strike_type_name} delta for trade at row {found_row}",
            "sheet_name": sheet_name,
            "trade_date": trade_date,
            "trade_time": trade_time,
            "sold_strike": sold_strike,
            "strike_type": strike_type_name,
            "strike_price": strike_price,
            "delta": delta,
            "delta_time": delta_time,
            "time_window": time_window,
            "updated_cell": cell_address,
            "found_row": found_row,
        }


async def close_trade_impl(
    url: str,
    file_name: str,
    sheet_name: str,
    trade_date: str,
    trade_time: str,
    close_date: str,
    close_time: str
) -> dict:
    """
    Core implementation for closing a trade by updating close date and time.
    Returns a dict with the result (not a JSON string).
    
    Args:
        url: SharePoint/OneDrive URL to the document library
        file_name: Excel file name with .xlsx extension
        sheet_name: Worksheet name
        trade_date: Date when the trade was opened (Column C)
        trade_time: Time when the trade was opened (Column E)
        close_date: Date when the trade was closed (Column F)
        close_time: Time when the trade was closed (Column G)
    
    Returns:
        Dictionary with operation result
    """
    logger.info(f"Closing trade: trade_date={trade_date}, trade_time={trade_time}, "
                f"close_date={close_date}, close_time={close_time}")
    
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
                "message": f"Worksheet '{sheet_name}' is empty",
            }
        
        # Step 2: Get columns C (date) and E (time) to find the matching row
        # Read both columns in one request for efficiency
        search_range = f"C1:E{row_count}"
        search_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{search_range}')"
        logger.info(f"Searching for trade with date='{trade_date}' and time='{trade_time}'")
        
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
                "message": f"Failed to read search columns: {error_message}",
                "status_code": search_response.status_code,
            }
        
        search_data = search_response.json()
        row_values = search_data.get("values", [])
        
        # Step 3: Find the row with matching date (C) and time (E)
        # Note: C is index 0, D is index 1, E is index 2
        found_row = None
        for i, row in enumerate(row_values):
            date_value = row[0] if len(row) > 0 else None  # Column C
            time_value = row[2] if len(row) > 2 else None  # Column E
            
            # Compare date
            date_match = compare_values_for_search(date_value, trade_date)
            
            # Compare time (normalize both values for comparison)
            time_match = False
            if time_value is not None and trade_time:
                # Normalize time values for comparison
                time_value_str = str(time_value).strip().upper()
                trade_time_str = trade_time.strip().upper()
                # Direct string comparison or parse and compare
                if time_value_str == trade_time_str:
                    time_match = True
                else:
                    # Try parsing both times and comparing
                    time_value_minutes = parse_time_to_minutes(str(time_value))
                    trade_time_minutes = parse_time_to_minutes(trade_time)
                    if time_value_minutes is not None and trade_time_minutes is not None:
                        time_match = (time_value_minutes == trade_time_minutes)
            
            if date_match and time_match:
                found_row = i + 1  # Excel rows are 1-indexed
                logger.info(f"Found matching trade at row {found_row}: date='{date_value}', time='{time_value}'")
                break
        
        if found_row is None:
            # Provide diagnostic info
            sample_rows = []
            for i, row in enumerate(row_values[:10]):
                date_val = row[0] if len(row) > 0 else "empty"
                time_val = row[2] if len(row) > 2 else "empty"
                sample_rows.append(f"Row {i+1}: date='{date_val}', time='{time_val}'")
            
            return {
                "status": "error",
                "message": f"No trade found with date '{trade_date}' and time '{trade_time}' in sheet '{sheet_name}'",
                "searched_rows": row_count,
                "sample_rows": sample_rows,
                "hint": "Ensure the trade_date and trade_time match exactly what's in the spreadsheet",
            }
        
        # Step 4: Update the close date (Column F) and close time (Column G)
        updated_cells = []
        errors = []
        
        # Update close date (Column F)
        close_date_address = f"F{found_row}"
        close_date_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{close_date_address}')"
        
        logger.info(f"Updating cell '{close_date_address}' with close_date '{close_date}'")
        
        close_date_response = await client.patch(
            close_date_url,
            headers=headers,
            json={"values": [[close_date]]},
            timeout=30.0,
        )
        
        if close_date_response.status_code == 200:
            updated_cells.append(close_date_address)
        else:
            error_data = close_date_response.json() if close_date_response.content else {}
            error_message = error_data.get("error", {}).get("message", close_date_response.text)
            errors.append({"cell": close_date_address, "error": error_message})
        
        # Update close time (Column G)
        close_time_address = f"G{found_row}"
        close_time_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{close_time_address}')"
        
        logger.info(f"Updating cell '{close_time_address}' with close_time '{close_time}'")
        
        close_time_response = await client.patch(
            close_time_url,
            headers=headers,
            json={"values": [[close_time]]},
            timeout=30.0,
        )
        
        if close_time_response.status_code == 200:
            updated_cells.append(close_time_address)
        else:
            error_data = close_time_response.json() if close_time_response.content else {}
            error_message = error_data.get("error", {}).get("message", close_time_response.text)
            errors.append({"cell": close_time_address, "error": error_message})
        
        if errors:
            return {
                "status": "partial_error",
                "message": "Some cells failed to update",
                "updated_cells": updated_cells,
                "errors": errors,
            }
        
        return {
            "status": "success",
            "message": f"Successfully closed trade at row {found_row}",
            "sheet_name": sheet_name,
            "trade_date": trade_date,
            "trade_time": trade_time,
            "close_date": close_date,
            "close_time": close_time,
            "updated_cells": updated_cells,
            "found_row": found_row,
        }


async def update_range_impl(
    url: str,
    file_name: str,
    sheet_name: str,
    address: str,
    values: list
) -> dict:
    """
    Core implementation for updating a range of cells in an Excel worksheet.
    Returns a dict with the result (not a JSON string).
    
    Args:
        url: SharePoint/OneDrive URL to the document library
        file_name: Excel file name with .xlsx extension
        sheet_name: Worksheet name
        address: Cell range address (e.g., "A1:C3")
        values: 2D list of values (already parsed from JSON)
    
    Returns:
        Dictionary with operation result
    """
    # Validate values is a 2D list
    if not isinstance(values, list) or not all(isinstance(row, list) for row in values):
        return {
            "status": "error",
            "message": "values must be a 2D array (list of lists)",
        }
    
    # Resolve URL to get drive_id, item_id, and site_id
    resolved = await resolve_excel_file_ids(url, file_name)
    if resolved.get("status") != "success":
        return resolved
    
    drive_id = resolved["drive_id"]
    item_id = resolved["item_id"]
    site_id = resolved.get("site_id")
    
    # Build the URL for range update
    workbook_url = build_workbook_url(drive_id, item_id, site_id)
    range_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"
    
    body = {"values": values}
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
            logger.info(f"Successfully updated range '{address}' in sheet '{sheet_name}'")
            return {
                "status": "success",
                "message": f"Successfully updated range '{address}' in sheet '{sheet_name}'",
                "file_name": resolved.get("file_name"),
                "sheet_name": sheet_name,
                "address": result_data.get("address", address),
                "row_count": result_data.get("rowCount"),
                "column_count": result_data.get("columnCount"),
            }
        else:
            error_data = response.json() if response.content else {}
            error_message = error_data.get("error", {}).get("message", response.text)
            return {
                "status": "error",
                "message": f"Failed to update range: {error_message}",
                "status_code": response.status_code,
            }


async def log_trades_impl(
    url: str,
    file_name: str,
    sheet_name: str,
    trades: list
) -> dict:
    """
    Core implementation for logging multiple trades to an Excel workbook.
    Returns a dict with the result (not a JSON string).
    
    Args:
        url: SharePoint/OneDrive URL to the document library
        file_name: Excel file name with .xlsx extension
        sheet_name: Worksheet name
        trades: List of trade dictionaries (already parsed from JSON)
    
    Returns:
        Dictionary with operation result
    """
    # Column configuration for trade tracker
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
    
    logger.info(f"excel.logTrades called with sheet_name='{sheet_name}'")
    
    if not isinstance(trades, list):
        return {
            "status": "error",
            "message": "trades must be a list of trade objects",
        }
    
    if len(trades) == 0:
        return {
            "status": "warning",
            "message": "No trades provided to log",
        }
    
    # Sort trades by open_date (ascending) then by open_time (ascending)
    def parse_trade_datetime(trade: dict) -> tuple:
        """Parse trade's open_date and open_time for sorting."""
        # Handle alternative field names
        open_date_str = trade.get("open_date") or trade.get("date") or trade.get("executed_date") or ""
        open_time_str = trade.get("open_time") or trade.get("time") or trade.get("executed_time") or ""
        
        date_key = datetime.max
        if open_date_str:
            parsed_date = parse_date_string(open_date_str)
            if parsed_date:
                date_key = parsed_date
        
        time_key = datetime.max.time()
        if open_time_str:
            time_formats = [
                "%I:%M %p", "%I:%M%p", "%H:%M", "%H:%M:%S", "%I:%M:%S %p",
            ]
            for fmt in time_formats:
                try:
                    parsed_time = datetime.strptime(open_time_str.strip().upper(), fmt)
                    time_key = parsed_time.time()
                    break
                except ValueError:
                    continue
        
        return (date_key, time_key)
    
    trades.sort(key=parse_trade_datetime)
    logger.info(f"Sorted {len(trades)} trades by open_date and open_time (ascending)")
    
    # Resolve URL to get drive_id, item_id, and site_id
    resolved = await resolve_excel_file_ids(url, file_name)
    if resolved.get("status") != "success":
        return {
            "status": "error",
            "message": f"Failed to resolve Excel file: {resolved.get('message')}",
        }
    
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
            return {
                "status": "error",
                "message": f"Failed to get worksheet data: {error_message}",
            }
        
        used_range_data = used_range_response.json()
        row_count = used_range_data.get("rowCount", 0)
        
        if row_count == 0:
            return {
                "status": "error",
                "message": f"Worksheet '{sheet_name}' is empty",
            }
        
        # Get column C values
        search_range = f"C1:C{row_count}"
        search_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{search_range}')"
        search_response = await client.get(search_url, headers=headers, timeout=30.0)
        
        if search_response.status_code != 200:
            error_data = search_response.json() if search_response.content else {}
            error_message = error_data.get("error", {}).get("message", search_response.text)
            return {
                "status": "error",
                "message": f"Failed to read column C: {error_message}",
            }
        
        search_data = search_response.json()
        column_values = search_data.get("values", [])
        
        # Find the last non-empty cell with a valid date value (searching from bottom)
        last_date_row = None
        
        for i in range(len(column_values) - 1, -1, -1):
            cell_value = column_values[i][0] if column_values[i] else None
            if cell_value is not None and cell_value != "":
                try:
                    if isinstance(cell_value, (int, float)) and 1 <= cell_value <= 2958465:
                        dt = excel_serial_to_date(cell_value)
                        last_date_row = i + 1
                        logger.info(f"Found last date at row {last_date_row}: Excel serial {cell_value} → {dt.strftime('%m/%d/%Y')}")
                        break
                    elif isinstance(cell_value, str):
                        parsed = parse_date_string(cell_value)
                        if parsed:
                            last_date_row = i + 1
                            logger.info(f"Found last date at row {last_date_row}: {cell_value}")
                            break
                except (ValueError, TypeError):
                    continue
        
        # Fallback: look for "Date of Purchase" header
        if last_date_row is None:
            logger.info("No valid dates found in column C, searching for 'Date of Purchase' header...")
            for i, row_value in enumerate(column_values):
                cell_value = row_value[0] if row_value else None
                if cell_value is not None and isinstance(cell_value, str):
                    if "date" in cell_value.lower() and "purchase" in cell_value.lower():
                        last_date_row = i + 1
                        logger.info(f"Found 'Date of Purchase' header at row {last_date_row}")
                        break
        
        if last_date_row is None:
            return {
                "status": "error",
                "message": f"Could not find any valid date or 'Date of Purchase' header in column C of sheet '{sheet_name}'",
            }
    
    logger.info(f"Logging {len(trades)} trades to {file_name}, sheet '{sheet_name}'")
    logger.info(f"Last date row: {last_date_row}, will start writing at row {last_date_row + 1}")
    
    results = []
    errors = []
    
    for i, trade in enumerate(trades):
        logger.info(f"Processing trade {i+1}: {json.dumps(trade)}")
        
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
        
        target_row = last_date_row + 1 + i
        
        logger.info(f"Logging trade {i+1}/{len(trades)}: strategy={mapped_strategy}, credit={trade.get('credit')}, contracts={trade.get('contracts')} to row {target_row}")
        
        updated_cells = []
        cell_errors = []
        
        async with httpx.AsyncClient() as write_client:
            for col_letter, field_name in COLUMN_MAP.items():
                value = values_dict.get(field_name, "")
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
        return {
            "status": "error",
            "message": f"All {len(trades)} trades failed to log",
            "errors": errors,
        }
    elif errors:
        return {
            "status": "partial_success",
            "message": f"Logged {len(results)} of {len(trades)} trades",
            "trades_logged": len(results),
            "trades_failed": len(errors),
            "file_name": file_name,
            "sheet_name": sheet_name,
            "start_row": last_date_row + 1,
            "results": results,
            "errors": errors,
        }
    else:
        return {
            "status": "success",
            "message": f"Successfully logged {len(results)} trades",
            "trades_logged": len(results),
            "file_name": file_name,
            "sheet_name": sheet_name,
            "start_row": last_date_row + 1,
            "results": results,
        }