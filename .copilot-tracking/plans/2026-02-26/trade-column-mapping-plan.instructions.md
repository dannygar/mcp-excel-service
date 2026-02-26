---
applyTo: '.copilot-tracking/changes/2026-02-26/trade-column-mapping-changes.md'
---
<!-- markdownlint-disable-file -->
# Implementation Plan: Trade Column Mapping Fix — Excel CSV to JSON Alignment

## Overview

Correct the COLUMN_MAP offset mismatch between the MCP server code and the actual Excel workbook layout, add the missing ATR field, implement grouped range writes for efficiency, update all dependent column references, and eliminate the duplicated inline logTrades logic in server.py.

## Objectives

### User Requirements

* Fix COLUMN_MAP to match the actual CSV/Excel workbook column layout (columns A–AC) — Source: research Discovery 1 (column offset mismatch)
* Add the missing ATR field to the write pipeline — Source: research Discovery 3 (ATR silently discarded)
* Update dependent mappings (closeTrade columns, delta columns, row search columns) — Source: research Dependent Mapping Updates section
* Document the complete column mapping between JSON input and Excel columns — Source: task request

### Derived Objectives

* Eliminate duplicated logTrades inline logic in server.py by delegating to `log_trades_impl()` — Derived from: research finding that server.py contains ~300 lines of near-identical inline code that would need to be maintained in two places
* Implement grouped range PATCH writes (Scenario A) to reduce API calls by ~45% — Derived from: research write strategy analysis showing 6 calls per trade vs. current 11
* Ensure formula columns (B, F, K, N, R) are never overwritten — Derived from: research Discovery 4 identifying 5 formula-dependent columns

## Context Summary

### Project Files

* mcp-server/core_operations.py (965 lines) — Primary implementation: COLUMN_MAP (L638–649), DELTA_COLUMN_MAPPING (L28–36), log_trades_impl (L617–965), close_trade_impl (L336–540)
* mcp-server/server.py (1430 lines) — MCP tools and REST endpoints: inline COLUMN_MAP (L434–454), inline write loop (L729–756), ~300 lines of duplicated logTrades logic (L480–780)
* mcp-server/config.py (222 lines) — Strategy mapping (no changes needed)
* mcp-server/excel_helpers.py (430 lines) — Helper functions (no changes needed)
* docs/samples/2026 GY Capital Group LLC Trade Tracker.csv — Reference CSV with 29 columns A–AC

### References

* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md — Complete column mapping research with offset analysis and corrected COLUMN_MAP
* .github/copilot-instructions.md — Repository conventions (FastMCP decorators, async/await, JSON returns)

### Standards References

* #file:../../.github/copilot-instructions.md — FastMCP decorators, async/await, JSON returns, logging conventions
* Python script instructions — Python 3.11+ conventions, logging, error handling

## Implementation Checklist

### [x] Implementation Phase 1: Eliminate server.py logTrades Duplication

<!-- parallelizable: false -->

* [x] Step 1.1: Refactor MCP tool `excel.logTrades` in server.py to delegate to `log_trades_impl()`
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 16–63)
* [x] Step 1.2: Verify REST endpoint `api_log_trades` delegates to `log_trades_impl()` correctly
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 65–95)
* [x] Step 1.3: Validate refactored delegation works (manual test with MCP Inspector)
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 97–113)

### [x] Implementation Phase 2: Correct COLUMN_MAP and Add ATR

<!-- parallelizable: false -->

* [x] Step 2.1: Update COLUMN_MAP in core_operations.py to corrected 12-column version
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 119–159)
* [x] Step 2.2: Add ATR field extraction to values_dict construction
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 161–189)
* [x] Step 2.3: Implement grouped range PATCH writes replacing per-cell loop
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 191–259)

### [x] Implementation Phase 3: Update Dependent Column References

<!-- parallelizable: true -->

* [x] Step 3.1: Update DELTA_COLUMN_MAPPING (shift all columns left by 1)
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 265–302)
* [x] Step 3.2: Update close_trade_impl column references (F→D, G→E, search C→A, search E→C)
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 304–352)
* [x] Step 3.3: Update row search column in log_trades_impl (C→A for date column search)
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 354–383)
* [x] Step 3.4: Update update_trade_with_delta_impl search range (C1:E → A1:C)
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 385–416)

### [x] Implementation Phase 4: Update server.py Inline References (if any remain)

<!-- parallelizable: true -->

* [x] Step 4.1: Remove any residual inline COLUMN_MAP or column references in server.py closeTrade tool
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 389–415)
* [x] Step 4.2: Verify updateTradeWithDelta MCP tool delegates properly and references correct columns
  * Details: .copilot-tracking/details/2026-02-26/trade-column-mapping-details.md (Lines 417–440)

### [x] Implementation Phase 5: Validation

<!-- parallelizable: false -->

* [x] Step 5.1: Run Python linting and syntax checks
  * Execute `uv run python -m py_compile server.py` and `uv run python -m py_compile core_operations.py`
  * Verify no import errors or syntax issues
* [x] Step 5.2: Verify column mapping correctness against CSV reference
  * Cross-reference corrected COLUMN_MAP against CSV header row (row 27)
  * Verify formula columns B, F, K, N, R are not in any write path
* [ ] Step 5.3: Manual integration test via MCP Inspector
  * Test logTrades with sample JSON input (test_logTrades.json)
  * Verify cells are written to correct columns in Excel workbook
  * Test closeTrade to verify updated column references (D, E)
  * Test updateTradeWithDelta to verify shifted delta columns
* [x] Step 5.4: Fix minor validation issues
  * Iterate on lint errors and build warnings
  * Apply fixes directly when corrections are straightforward
* [x] Step 5.5: Report blocking issues
  * Document issues requiring additional research
  * Provide next steps and recommended planning if large-scale fixes are needed

## Planning Log

See [trade-column-mapping-log.md](.copilot-tracking/plans/logs/2026-02-26/trade-column-mapping-log.md) for discrepancy tracking, implementation paths considered, and suggested follow-on work.

## Dependencies

* Python 3.11+ with `uv` package manager
* `httpx` for async HTTP calls (existing dependency)
* `fastmcp` for MCP server decorators (existing dependency)
* Microsoft Graph API access with Files.ReadWrite.All permissions (existing)
* MCP Inspector (`yarn inspector`) for manual testing
* Access to the target Excel workbook for integration testing

## Success Criteria

* Corrected COLUMN_MAP with 12 fields maps JSON input to columns A, C, G–J, L–M, O–Q, S — Traces to: research Discovery 5
* ATR field is extracted from trade input and written to column O — Traces to: research Discovery 3
* Grouped range writes reduce API calls from 11 to 6 per trade — Traces to: research Scenario A analysis
* closeTrade writes to columns D/E and searches columns A/C — Traces to: research Dependent Mapping Updates
* DELTA_COLUMN_MAPPING references columns T–AA (shifted left by 1) — Traces to: research Dependent Mapping Updates
* Formula columns (B, F, K, N, R) are never written to — Traces to: research Discovery 4
* server.py logTrades delegates to core_operations.log_trades_impl() with no inline duplication — Traces to: derived objective (code maintainability)
* All changes pass Python syntax checks with no errors — Traces to: standard validation
