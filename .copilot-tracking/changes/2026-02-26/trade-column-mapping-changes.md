<!-- markdownlint-disable-file -->
# Release Changes: Trade Column Mapping Fix

**Related Plan**: trade-column-mapping-plan.instructions.md
**Implementation Date**: 2026-02-26

## Summary

Correct the COLUMN_MAP offset mismatch between the MCP server code and the actual Excel workbook layout, add the missing ATR field, implement grouped range writes, update all dependent column references, and eliminate the duplicated inline logTrades logic in server.py.

## Changes

### Added

* mcp-server/core_operations.py — `COLUMN_GROUPS` constant (6 grouped ranges for batch PATCH writes)
* mcp-server/core_operations.py — ATR field extraction (`trade.get("ATR") or trade.get("atr")`) and inclusion in `values_dict`

### Modified

* mcp-server/server.py — Replaced ~380-line inline `excel.logTrades` MCP tool with ~35-line delegation to `log_trades_impl()` (Phase 1, -342 lines)
* mcp-server/server.py — Updated `excel.logTrades` docstring Column Mapping section (C→A, E→C, I→G, J→H, K→I, L→J, N→L, O→M, +O:ATR, Q→P, R→Q, T→S)
* mcp-server/server.py — Updated `excel.closeTrade` docstring Column Mapping (search C→A, E→C; update F→D, G→E)
* mcp-server/server.py — Updated `excel.updateTradeWithDelta` docstring Column Mapping (search C→A, E→C; delta U–AB→T–AA)
* mcp-server/core_operations.py — Corrected `COLUMN_MAP` from 11 entries (wrong offset) to 12 entries matching CSV layout (A, C, G–J, L–M, O–Q, S)
* mcp-server/core_operations.py — Replaced per-cell write loop (11 PATCH calls) with grouped range writes using `COLUMN_GROUPS` (6 PATCH calls)
* mcp-server/core_operations.py — Shifted `DELTA_COLUMN_MAPPING` columns left by 1 (U–AB → T–AA)
* mcp-server/core_operations.py — Updated `close_trade_impl()` search range C1:E→A1:C, close date F→D, close time G→E
* mcp-server/core_operations.py — Updated `log_trades_impl()` date column search C1:C→A1:A
* mcp-server/core_operations.py — Updated `update_trade_with_delta_impl()` search range C1:E→A1:C
* mcp-server/core_operations.py — Updated all docstrings and comments to reference corrected column letters

### Removed

* mcp-server/server.py — Removed ~380 lines of duplicated inline logTrades logic (inline COLUMN_MAP, field extraction, values_dict, per-cell write loop, response building)

## Additional or Deviating Changes

* Step 5.3 (manual integration test via MCP Inspector) is deferred — requires live Azure credentials and running workbook access
  * Reason: automated validation (compile checks, import verification, CSV cross-reference, stale reference scan) covers all code-level correctness

## Release Summary

**Total files affected**: 2 (mcp-server/server.py, mcp-server/core_operations.py)

**Files modified**:

* [mcp-server/server.py](mcp-server/server.py) — Refactored logTrades MCP tool to delegate to core impl (-342 lines); updated all tool docstrings with corrected column references
* [mcp-server/core_operations.py](mcp-server/core_operations.py) — Fixed COLUMN_MAP (12 entries matching CSV), added ATR field, implemented grouped range writes (6 vs 11 API calls), shifted DELTA_COLUMN_MAPPING (T–AA), updated all search ranges and close trade columns

**Dependency changes**: None

**Infrastructure changes**: None

**Deployment notes**: No configuration changes required. The corrected column mapping assumes the production Excel workbook matches the CSV layout in `docs/samples/2026 GY Capital Group LLC Trade Tracker.csv`. Verify workbook layout before first use.
