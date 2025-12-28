"""
Core Excel Operations Module

Provides the core implementation functions for Excel operations.
These functions are called by the MCP tools in server.py.
"""

import logging

import httpx

from graph_api import get_graph_headers, build_workbook_url
from excel_helpers import resolve_excel_file_ids, compare_values_for_search

logger = logging.getLogger("mcp-excel-server")


# =============================================================================
# Core Implementation Functions (called by tools)
# =============================================================================

async def update_row_by_lookup_impl(
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
    
    Args:
        url: SharePoint/OneDrive URL to the document library
        file_name: Excel file name with .xlsx extension
        sheet_name: Worksheet name
        search_column: Column letter to search
        reference_value: Value to find in the search column
        target_columns_list: List of column letters to update
        values_list: List of values to write
        row_offset: Number of rows below the found row to update
    
    Returns:
        Dictionary with operation result
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
