<!-- markdownlint-disable-file -->
# Implementation Details: Trade Column Mapping Fix — Excel CSV to JSON Alignment

## Context Reference

Sources: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md, mcp-server/core_operations.py (965 lines), mcp-server/server.py (1430 lines)

## Implementation Phase 1: Eliminate server.py logTrades Duplication

<!-- parallelizable: false -->

### Step 1.1: Refactor MCP tool `excel.logTrades` to delegate to `log_trades_impl()`

Replace the ~300 lines of inline logTrades logic in server.py (lines 434–780) with a call to `log_trades_impl()` from core_operations.py, matching the delegation pattern used by closeTrade (line 253) and updateRange.

The MCP tool function should:
1. Parse the `trades` JSON string into `trade_results` list (keep existing parsing at lines 388–430)
2. Call `log_trades_impl(url, file_name, sheet_name, trade_results)` and return the result
3. Remove all inline logic: COLUMN_MAP, field extraction, values_dict, write loop, response building

Files:
* mcp-server/server.py — Remove lines ~434–780 (inline COLUMN_MAP, field extraction, values_dict, write loop), replace with delegation call
* mcp-server/core_operations.py — No changes (already has `log_trades_impl`)

Discrepancy references:
* Addresses DD-01 (inline duplication as maintenance risk)

Success criteria:
* MCP tool `excel.logTrades` function body is ~20 lines (parse JSON + delegate + return)
* No COLUMN_MAP definition in server.py
* No field extraction or values_dict logic in server.py

Context references:
* mcp-server/server.py (Lines 382–780) — Current inline logTrades implementation
* mcp-server/server.py (Lines 205–260) — closeTrade delegation pattern to follow
* mcp-server/core_operations.py (Lines 617–965) — Target `log_trades_impl()` function

Dependencies:
* None — first step in the plan

### Step 1.2: Verify REST endpoint `api_log_trades` delegates to `log_trades_impl()` correctly

Verify (do not refactor) that the REST endpoint at server.py line 920 already calls `log_trades_impl()` directly. The subagent report indicates the REST endpoint may already delegate correctly. If it calls the MCP tool wrapper instead, refactor to call `log_trades_impl()` directly matching the closeTrade REST endpoint pattern.

Files:
* mcp-server/server.py — Update `api_log_trades` function (around lines 920–960) to delegate to `log_trades_impl()`

Discrepancy references:
* Addresses DD-01 (REST endpoint should use same delegation pattern as closeTrade)

Success criteria:
* REST endpoint calls `log_trades_impl()` directly, not the MCP tool wrapper
* Consistent delegation pattern across all REST endpoints

Context references:
* mcp-server/server.py (Lines 920–960) — Current api_log_trades endpoint
* mcp-server/server.py (Lines 1063–1127) — closeTrade REST endpoint pattern to follow

Dependencies:
* Step 1.1 completion (MCP tool must be refactored first to avoid circular call)

### Step 1.3: Validate refactored delegation works

Verify the refactored server.py compiles without errors and that the MCP tool and REST endpoint both route to `log_trades_impl()`.

Validation commands:
* `cd mcp-server && uv run python -m py_compile server.py` — Syntax check
* `cd mcp-server && uv run python -c "from server import *"` — Import check
* Manual: Start server (`uv run python server.py`), call logTrades via MCP Inspector

Files:
* No file changes — validation only

Success criteria:
* `py_compile server.py` exits with code 0
* Server starts without import errors
* logTrades tool appears in MCP Inspector tool list

Dependencies:
* Steps 1.1 and 1.2 completion

## Implementation Phase 2: Correct COLUMN_MAP and Add ATR

<!-- parallelizable: false -->

### Step 2.1: Update COLUMN_MAP in core_operations.py to corrected 12-column version

Replace the current 11-entry COLUMN_MAP (lines 638–649) with the corrected 12-entry version that matches the CSV workbook layout. This is the primary fix for Discovery 1 (column offset mismatch).

Current COLUMN_MAP (wrong — shifted +2/+1):
```python
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
```

Corrected COLUMN_MAP (matches CSV layout):
```python
COLUMN_MAP = {
    "A": "open_date",         # Date of Purchase
    "C": "open_time",         # Time IN
    "G": "strategy",          # Strategy (after map_strategy_name())
    "H": "credit",            # Credit (per-contract, after ÷ contracts × 100)
    "I": "debit",             # Debit (per-contract, after ÷ contracts × 100)
    "J": "contracts",         # Contracts
    "L": "open_fees",         # Open Fees
    "M": "close_fees",        # Close Fees
    "O": "ATR",               # ATR — NEW FIELD
    "P": "sold_call_strike",  # Short Calls
    "Q": "sold_put_strike",   # Short Puts
    "S": "width",             # Width
}
```

Files:
* mcp-server/core_operations.py — Replace COLUMN_MAP at lines 638–649

Discrepancy references:
* Addresses DR-01 (column offset mismatch — critical discovery)

Success criteria:
* COLUMN_MAP has 12 entries (was 11)
* Column letters match CSV positions: A, C, G, H, I, J, L, M, O, P, Q, S
* No formula columns (B, F, K, N, R) appear in the map
* New "O": "ATR" entry present

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Discovery 5) — Corrected mapping with evidence
* docs/samples/2026 GY Capital Group LLC Trade Tracker.csv (Row 27) — Column headers

Dependencies:
* Phase 1 completion (server.py no longer has its own COLUMN_MAP)

### Step 2.2: Add ATR field extraction to values_dict construction

Add ATR field extraction to the `values_dict` construction block in `log_trades_impl()` (around lines 874–886). The ATR value is a float passed directly from the JSON input — no transformation needed.

Add after the existing field extractions (around line 872):
```python
# ATR
atr_value = trade.get("ATR") or trade.get("atr")
```

And add to `values_dict`:
```python
values_dict = {
    "open_date": open_date,
    "open_time": open_time,
    "strategy": strategy_name,
    "credit": credit_value,
    "debit": debit_value,
    "contracts": contracts,
    "open_fees": open_fees,
    "close_fees": close_fees,
    "ATR": atr_value,               # NEW
    "sold_call_strike": sold_call,
    "sold_put_strike": sold_put,
    "width": width_value,
}
```

Files:
* mcp-server/core_operations.py — Add ATR extraction (after line ~872) and update values_dict (lines 874–886)

Discrepancy references:
* Addresses DR-02 (ATR field silently discarded)

Success criteria:
* `values_dict` contains 12 keys (was 11)
* ATR value is extracted from trade dict with case-insensitive fallback ("ATR" or "atr")
* ATR value is written as-is (float, no transformation)

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Discovery 3) — ATR missing from COLUMN_MAP
* mcp-server/core_operations.py (Lines 874–886) — Current values_dict

Dependencies:
* Step 2.1 completion (COLUMN_MAP must include "O": "ATR")

### Step 2.3: Implement grouped range PATCH writes replacing per-cell loop

Replace the per-cell write loop (lines 894–919) with a grouped range write approach. Define `COLUMN_GROUPS` that map contiguous column ranges to field lists, then write each group as a single PATCH call.

Define COLUMN_GROUPS (place near COLUMN_MAP, after line 649):
```python
COLUMN_GROUPS = [
    ("A", ["open_date"]),
    ("C", ["open_time"]),
    ("G", ["strategy", "credit", "debit", "contracts"]),
    ("L", ["open_fees", "close_fees"]),
    ("O", ["ATR", "sold_call_strike", "sold_put_strike"]),
    ("S", ["width"]),
]
```

Replace the per-cell write loop with grouped range logic:
```python
for start_col, fields in COLUMN_GROUPS:
    values = [values_dict.get(f, "") for f in fields]

    # Skip group if all values are empty/None
    if all(v == "" or v is None for v in values):
        continue

    # Replace None with empty string for partial groups
    values = ["" if v is None else v for v in values]

    if len(fields) == 1:
        address = f"{start_col}{row_number}"
    else:
        end_col = chr(ord(start_col) + len(fields) - 1)
        address = f"{start_col}{row_number}:{end_col}{row_number}"

    url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"

    try:
        resp = await client.patch(
            url, headers=headers,
            json={"values": [values]},
            timeout=30.0
        )
        if resp.status_code in (200, 201):
            updated_cells.append(address)
            logging.info(f"Updated range {address}")
        else:
            errors.append(f"Failed to update {address}: {resp.text}")
    except Exception as e:
        errors.append(f"Error updating {address}: {str(e)}")
```

Files:
* mcp-server/core_operations.py — Add COLUMN_GROUPS (after line ~649), replace write loop (lines 894–919)

Discrepancy references:
* Addresses selected implementation path (Scenario A — grouped range writes)

Success criteria:
* 6 PATCH calls per trade (down from 11)
* Contiguous column groups match the CSV layout
* Empty/None values handled correctly (skip all-empty groups, write "" for partial)
* Updated cells tracked for response reporting

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Scenario A) — Grouped range write design
* mcp-server/core_operations.py (Lines 894–919) — Current per-cell write loop to replace

Dependencies:
* Steps 2.1 and 2.2 completion (COLUMN_MAP and ATR must be updated first)

## Implementation Phase 3: Update Dependent Column References

<!-- parallelizable: true -->

### Step 3.1: Update DELTA_COLUMN_MAPPING (shift all columns left by 1)

Update `DELTA_COLUMN_MAPPING` (lines 28–36) to reflect the corrected column positions. All delta columns shift left by 1 position.

Current mapping:
```python
DELTA_COLUMN_MAPPING = {
    ("9:30-11:00", "C"): "U",   ("9:30-11:00", "P"): "V",
    ("11:00-12:00", "C"): "W",  ("11:00-12:00", "P"): "X",
    ("1:00-2:00", "C"): "Y",    ("1:00-2:00", "P"): "Z",
    ("2:30-3:30", "C"): "AA",   ("2:30-3:30", "P"): "AB",
}
```

Corrected mapping (matches CSV columns T–AA):
```python
DELTA_COLUMN_MAPPING = {
    ("9:30-11:00", "C"): "T",   ("9:30-11:00", "P"): "U",
    ("11:00-12:00", "C"): "V",  ("11:00-12:00", "P"): "W",
    ("1:00-2:00", "C"): "X",    ("1:00-2:00", "P"): "Y",
    ("2:30-3:30", "C"): "Z",    ("2:30-3:30", "P"): "AA",
}
```

Also update `TIME_WINDOWS` if the time labels reference column positions (review lines 39–45).

**Delta column header alignment verification:**

| Time Window | Strike | Old Column | New Column | CSV Header |
|---|---|:---:|:---:|---|
| 9:30–11:00 | Call | U | T | C Delta at 10am |
| 9:30–11:00 | Put | V | U | P Delta at 10am |
| 11:00–12:00 | Call | W | V | C Delta at 11:30 am |
| 11:00–12:00 | Put | X | W | P Delta at 11:30am |
| 1:00–2:00 | Call | Y | X | C Delta at 1:30pm |
| 1:00–2:00 | Put | Z | Y | P Delta at 1:30pm |
| 2:30–3:30 | Call | AA | Z | C Delta at 3:00pm |
| 2:30–3:30 | Put | AB | AA | P Delta at 3:00pm |

Files:
* mcp-server/core_operations.py — Update DELTA_COLUMN_MAPPING (lines 28–36)

Discrepancy references:
* Addresses dependent mapping update (research section: Dependent Mapping Updates → updateTradeWithDelta)

Success criteria:
* Delta columns reference T–AA (was U–AB)
* Each (time_window, strike_type) tuple maps to the correct CSV column
* Column letters verified against CSV header names

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Dependent Mapping Updates → updateTradeWithDelta) — Column shift table
* mcp-server/core_operations.py (Lines 28–45) — Current DELTA_COLUMN_MAPPING and TIME_WINDOWS

Dependencies:
* Phase 2 completion (COLUMN_MAP corrected first to maintain consistency)

### Step 3.2: Update close_trade_impl column references

Update all hardcoded column references in `close_trade_impl()` (lines 336–540) to match the CSV layout.

Changes required:

1. **Search range**: Change `C1:E{row_count}` to `A1:C{row_count}` (lines 414–419)
   - Column A (was C): Date of Purchase (search by open date)
   - Column C (was E): Time IN (search by open time)

2. **Column index offsets in row matching** (lines 430–458):
   - Date column index: 0 (column A, first in the read range) — was already index 0 in `C1:E` → remains index 0 in `A1:C`
   - Time column index: 2 (column C, third in the read range) — was already index 2 in `C1:E` → remains index 2 in `A1:C`

3. **Close date address** (line 477):
   - Change `f"F{found_row}"` to `f"D{found_row}"` (column D = Close Day in CSV)

4. **Close time address** (line 501):
   - Change `f"G{found_row}"` to `f"E{found_row}"` (column E = Time OUT in CSV)

Files:
* mcp-server/core_operations.py — Update close_trade_impl() at lines 414–419, 477, 501

Discrepancy references:
* Addresses dependent mapping update (research section: Dependent Mapping Updates → closeTrade)

Success criteria:
* Search range reads columns A:C (was C:E)
* Close date written to column D (was F)
* Close time written to column E (was G)
* Row matching logic still works (index 0 = date, index 2 = time within the read range)

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Dependent Mapping Updates → closeTrade) — Column shift table
* mcp-server/core_operations.py (Lines 336–540) — Current close_trade_impl

Dependencies:
* Phase 2 completion (COLUMN_MAP corrected first)

### Step 3.3: Update row search column in log_trades_impl (C→A for date column search)

Update the date column search in `log_trades_impl()` (around lines 750–800) that reads column C to find the last date row for determining where to append new trades. Change to read column A instead.

Search for any references to column "C" used for date-based row lookup in the log_trades_impl function and update to column "A".

This includes:
* The used range read that determines existing rows (around lines 760–770)
* The bottom-up search for the last populated date cell (around lines 770–790)
* Any address construction referencing column C for date lookup

Files:
* mcp-server/core_operations.py — Update column references in log_trades_impl date search section (lines ~750–800)

Discrepancy references:
* Addresses dependent mapping update (research section: Row Search Column Updates)

Success criteria:
* Date column search reads column A (was C)
* Bottom-up row search still correctly identifies the last populated date row
* New trade rows are appended in the correct position

Context references:
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Row Search Column Updates) — Column A is Date of Purchase
* mcp-server/core_operations.py (Lines 750–800) — Current date column search logic

Dependencies:
* Step 2.1 completion (COLUMN_MAP must be corrected before search references)

### Step 3.4: Update update_trade_with_delta_impl search range (C1:E → A1:C)

Update the search range in `update_trade_with_delta_impl()` (core_operations.py line 221) from `C1:E{row_count}` to `A1:C{row_count}`. This function searches for a trade by date and time to determine which row to write delta values to.

Changes required:

1. **Search range** (line 221): Change `f"C1:E{row_count}"` to `f"A1:C{row_count}"`
2. **Column index comments** (line 237): Update comments from "Column C" / "Column E" to "Column A" / "Column C"
3. **Docstring** (lines 131–136): Update parameter comments from "Column C" / "Column E" to "Column A" / "Column C"

The row matching logic uses index 0 (date) and index 2 (time) within the read range. Since the range shifts from C:E to A:C, the relative indices remain the same:
* Index 0: Column A (was C) — Date of Purchase
* Index 2: Column C (was E) — Time IN

Files:
* mcp-server/core_operations.py — Update search range at line 221, comments at line 237, docstring at lines 131–136

Discrepancy references:
* Addresses VF-01 (validator finding — update_trade_with_delta_impl search range not covered)

Success criteria:
* Search range reads columns A:C (was C:E)
* Row matching logic still works (index 0 = date, index 2 = time)
* Docstring references corrected columns

Context references:
* mcp-server/core_operations.py (Lines 117–270) — update_trade_with_delta_impl function
* .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Dependent Mapping Updates) — Column shift applies to all search ranges

Dependencies:
* Phase 2 completion (COLUMN_MAP corrected first)

## Implementation Phase 4: Update server.py Inline References (if any remain)

<!-- parallelizable: true -->

### Step 4.1: Remove any residual inline COLUMN_MAP or column references in server.py closeTrade tool

Verify that the closeTrade MCP tool (server.py lines 205–277) delegates entirely to `close_trade_impl()` and contains no hardcoded column references. If it does delegate cleanly (as the subagent report confirms), no changes are needed.

If any column letters are hardcoded in the MCP tool (e.g., for parameter documentation or docstring examples), update them to match the corrected layout.

Files:
* mcp-server/server.py — Review lines 205–277 for any hardcoded column references

Discrepancy references:
* Preventive check — ensure no column references leak through delegation boundary

Success criteria:
* closeTrade MCP tool contains no hardcoded column letters
* closeTrade delegates entirely to close_trade_impl()

Context references:
* mcp-server/server.py (Lines 205–260) — closeTrade MCP tool
* mcp-server/server.py (Lines 1063–1127) — closeTrade REST endpoint

Dependencies:
* Phase 3 completion (close_trade_impl already updated)

### Step 4.2: Verify updateTradeWithDelta MCP tool delegates properly and references correct columns

Check the updateTradeWithDelta MCP tool in server.py for any inline column references. If it delegates to an impl function that uses DELTA_COLUMN_MAPPING, verify the delegation is clean.

If the MCP tool has inline delta column references, update them to match the corrected T–AA layout (was U–AB).

Files:
* mcp-server/server.py — Review updateTradeWithDelta MCP tool for inline column references

Discrepancy references:
* Preventive check — ensure delta column references are consistent

Success criteria:
* updateTradeWithDelta has no inline column letter references (or they match T–AA)
* Delegation to impl function is clean

Context references:
* mcp-server/core_operations.py (Lines 28–36) — Updated DELTA_COLUMN_MAPPING

Dependencies:
* Step 3.1 completion (DELTA_COLUMN_MAPPING updated)

## Implementation Phase 5: Validation

<!-- parallelizable: false -->

### Step 5.1: Run Python linting and syntax checks

Execute all validation commands for the project:
* `cd mcp-server && uv run python -m py_compile server.py`
* `cd mcp-server && uv run python -m py_compile core_operations.py`
* `cd mcp-server && uv run python -m py_compile config.py`
* `cd mcp-server && uv run python -m py_compile excel_helpers.py`
* `cd mcp-server && uv run python -c "from core_operations import log_trades_impl, close_trade_impl; print('imports OK')"`

### Step 5.2: Verify column mapping correctness against CSV reference

Manual verification checklist:
1. Open `docs/samples/2026 GY Capital Group LLC Trade Tracker.csv` row 27 (headers)
2. Verify each COLUMN_MAP entry matches the header at that column position
3. Verify formula columns B, F, K, N, R are NOT in COLUMN_MAP
4. Verify DELTA_COLUMN_MAPPING columns T–AA match CSV headers for delta columns
5. Verify close_trade_impl writes to D (Close Day) and E (Time OUT)

### Step 5.3: Manual integration test via MCP Inspector

Test workflow:
1. Start server: `cd mcp-server && uv run python server.py`
2. Launch MCP Inspector: `yarn inspector`
3. Call `excel.logTrades` with sample trade data from test_logTrades.json
4. Verify cells written to correct columns (A, C, G–J, L–M, O–Q, S)
5. Call `excel.closeTrade` — verify writes to columns D and E
6. Call `excel.updateTradeWithDelta` — verify writes to columns T–AA

### Step 5.4: Fix minor validation issues

Iterate on lint errors, build warnings, and test failures. Apply fixes directly when corrections are straightforward and isolated.

### Step 5.5: Report blocking issues

When validation failures require changes beyond minor fixes:
* Document the issues and affected files.
* Provide the user with next steps.
* Recommend additional research and planning rather than inline fixes.
* Avoid large-scale refactoring within this phase.

## Dependencies

* Python 3.11+ with `uv` package manager
* Microsoft Graph API access for integration testing
* MCP Inspector for manual testing

## Success Criteria

* All Python files compile without errors
* COLUMN_MAP has 12 entries matching CSV column positions
* Grouped writes produce 6 PATCH calls per trade
* closeTrade writes to D/E, searches A/C
* DELTA_COLUMN_MAPPING maps to T–AA
* No inline logTrades logic remains in server.py
* Formula columns (B, F, K, N, R) are never in any write path
