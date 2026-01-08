"""
Core Excel Operations Module

Provides the core implementation functions for Excel operations.
These functions are called by the MCP tools in server.py.
"""

import logging
from datetime import datetime

import httpx

from graph_api import get_graph_headers, build_workbook_url
from excel_helpers import resolve_excel_file_ids, compare_values_for_search

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
