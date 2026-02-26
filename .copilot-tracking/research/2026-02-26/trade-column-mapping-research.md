<!-- markdownlint-disable-file -->
# Task Research: Trade Column Mapping — Excel CSV to JSON Input

Analyze how the `logTrades` JSON input structure maps to Excel spreadsheet columns in the Trade Tracker workbook, and determine the optimal column-update strategy.

## Task Implementation Requests

* Map each field in the `trade_results[]` JSON input to the corresponding Excel column in the ORB TRANSACTION TRACKER table
* Identify columns that are written vs. left empty vs. computed (formula)
* Determine the best Graph API write strategy (per-cell, per-row, grouped ranges)
* Document any gaps between the JSON input and the Excel schema
* Document the critical column offset mismatch between the current code and the CSV layout

## Scope and Success Criteria

* Scope: Column mapping between `logTrades` JSON input and Trade Tracker Excel columns A–AC (row 27 headers). Includes current `log_trades_impl()` code analysis and offset verification.
* Assumptions:
  * The attached CSV (`docs/samples/2026 GY Capital Group LLC Trade Tracker.csv`) is representative of the actual `.xlsx` workbook column layout
  * Row 27 (1-indexed) contains column headers in the Excel workbook
  * The `excel_row` field in JSON indicates the target row number
  * The actual `.xlsx` workbook targeted by the code may have a different column layout (2 extra leading columns)
* Success Criteria:
  * Complete mapping table: JSON field → Excel column letter + header name for the CSV layout
  * Identification of columns not populated by logTrades (left for closeTrade, updateDelta, or manual entry)
  * Recommended write strategy with evidence from codebase and Graph API constraints
  * Quantified column offset between current code and CSV
  * Corrected COLUMN_MAP for the CSV layout

## Outline

1. JSON Input Field Inventory
2. CSV Column Schema (from CSV analysis)
3. Current Code Mapping (from `log_trades_impl()`)
4. Critical Discovery: Column Offset Mismatch
5. Corrected Column Mapping
6. Formula Columns (must not overwrite)
7. Write Strategy Analysis (3 scenarios)
8. Dependent Mapping Updates (closeTrade, updateDelta)
9. Recommendations

## Research Executed

### File Analysis

* [mcp-server/core_operations.py](mcp-server/core_operations.py)
  * COLUMN_MAP at lines 637–649: 11 entries mapping columns C, E, I, J, K, L, N, O, Q, R, T
  * Per-cell write loop at lines 893–916: iterates COLUMN_MAP, individual PATCH per non-empty field
  * DELTA_COLUMN_MAPPING at lines 27–36: U/V (9:30–11), W/X (11–12), Y/Z (1–2pm), AA/AB (2:30–3:30)
  * `close_trade_impl()` writes columns F (close_date) and G (close_time), searches C and E for row lookup
  * `values_dict` construction at lines 802–893: 11 fields, no ATR field
* [mcp-server/server.py](mcp-server/server.py)
  * Identical COLUMN_MAP at lines 441–452 (duplicated inline, not delegating to `log_trades_impl()`)
  * MCP tool `excel.logTrades` spans lines 382–815 with full inline logic
  * REST endpoint `api_log_trades` at line 920 correctly delegates to `log_trades_impl()`
* [mcp-server/config.py](mcp-server/config.py)
  * `map_strategy_name()` at lines 189–222 with 50+ strategy mappings
  * `STRATEGY_MAPPING` dict and `STRATEGY_KEYWORDS` pattern matching
* [docs/samples/2026 GY Capital Group LLC Trade Tracker.csv](docs/samples/2026%20GY%20Capital%20Group%20LLC%20Trade%20Tracker.csv)
  * 99 lines total, header row at row 27, 29 data rows (28–56)
  * 29 columns A–AC (Date of Purchase through Notes)
  * 3 strategies: VCCS (11 trades), VPCS (17 trades), IC (1 trade)

### Code Search Results

* `COLUMN_MAP`: Identical definition in both `server.py:441` and `core_operations.py:637`
* `DELTA_COLUMN_MAPPING`: Defined only in `core_operations.py:27–36`
* `close_trade_impl`: Writes to columns F and G, searches columns C and E
* `map_strategy_name`: Located in `config.py:189–222`

### External Research

* Graph API `PATCH /range(address=...)` with `{"values": [[...]]}` body — supports single-cell, single-row, or multi-cell range writes
* Graph API batch endpoint: `POST /$batch` — supports up to 20 requests per batch; reduces HTTP round-trips

### Project Conventions

* Standards referenced: `.github/copilot-instructions.md` — FastMCP decorators, async/await, JSON returns
* Instructions followed: Return JSON strings from tools, use `logging` module, keep secrets in env vars

## Key Discoveries

### Discovery 1: Column Offset Mismatch (Critical)

The current code's `COLUMN_MAP` is shifted relative to the CSV column layout. Verified by matching each JSON field value against its position in the CSV data for the 2/26/2026 VCCS trade.

**Evidence — 2/26/2026 VCCS trade (order_id 442689429):**

| JSON Field | JSON Value | CSV Column | CSV Position | Code Column | Offset |
|---|---|---|---|---|---|
| `open_date` | `02/26/2026` | A (Date of Purchase) | 0 | C | +2 |
| `open_time` | `10:11:10` | C (Time IN) → "10:11 AM" | 2 | E | +2 |
| `strategy` | → `VCCS` | G (Strategy) | 6 | I | +2 |
| `credit` | `375.0 / (25×100)` = `$0.15` | H (Credit) | 7 | J | +2 |
| `debit` | `0.0` (skipped) | I (Debit) | 8 | K | +2 |
| `contracts` | `25` | J (Contracts) | 9 | L | +2 |
| `open_fees` | `$86.24` | L (Open Fees) | 11 | N | +2 |
| `close_fees` | `null` (skipped) | M (Close Fees) | 12 | O | +2 |
| `sold_call_strike` | `6995` → `$6,995` | P (Short Calls) | 15 | Q | +1 |
| `sold_put_strike` | `null` (skipped) | Q (Short Puts) | 16 | R | +1 |
| `width` | `15` → `$15` | S (Width) | 18 | T | +1 |

**Pattern:** Offset is +2 for the first 8 fields (A→C through M→O), then drops to +1 for the last 3 fields (P→Q through S→T).

**Root cause:** The code was built for a workbook with:
1. Two extra leading columns (A–B) before the trade data
2. No ATR column between Profit and Short Calls

The CSV workbook starts trade data at column A and includes an ATR column at position O, which shifts the post-ATR columns by one position relative to the code's expectations.

### Discovery 2: Reconstructed Workbook Layouts

**Actual CSV layout (column A = position 0):**

| Position | Column | Header | Type |
|:---:|:---:|---|---|
| 0 | A | Date of Purchase | Data (open_date) |
| 1 | B | Open Day | Formula |
| 2 | C | Time IN | Data (open_time) |
| 3 | D | Close Day | Data (close_date — closeTrade) |
| 4 | E | Time OUT | Data (close_time — closeTrade) |
| 5 | F | Trade Time | Formula |
| 6 | G | Strategy | Data (strategy) |
| 7 | H | Credit | Data (credit per-contract) |
| 8 | I | Debit | Data (debit per-contract) |
| 9 | J | Contracts | Data (contracts) |
| 10 | K | Net Profit | Formula |
| 11 | L | Open Fees | Data (open_fees) |
| 12 | M | Close Fees | Data (close_fees) |
| 13 | N | Profit | Formula |
| 14 | O | ATR | Data (ATR) |
| 15 | P | Short Calls | Data (sold_call_strike) |
| 16 | Q | Short Puts | Data (sold_put_strike) |
| 17 | R | Spread Width | Formula |
| 18 | S | Width | Data (width) |
| 19 | T | C Delta at 10am | Data (delta) |
| 20 | U | P Delta at 10am | Data (delta) |
| 21 | V | C Delta at 11:30 am | Data (delta) |
| 22 | W | P Delta at 11:30am | Data (delta) |
| 23 | X | C Delta at 1:30pm | Data (delta) |
| 24 | Y | P Delta at 1:30pm | Data (delta) |
| 25 | Z | C Delta at 3:00pm | Data (delta) |
| 26 | AA | P Delta at 3:00pm | Data (delta) |
| 27 | AB | Rolls | Data (manual) |
| 28 | AC | Notes | Data (manual) |

**Code's assumed workbook layout (what the current COLUMN_MAP targets):**

| Position | Column | Inferred Header | Type |
|:---:|:---:|---|---|
| 0 | A | (extra/label) | Unknown |
| 1 | B | (extra/label) | Unknown |
| 2 | C | Date of Purchase | Data (open_date) |
| 3 | D | Open Day | Formula |
| 4 | E | Time IN | Data (open_time) |
| 5 | F | Close Day | Data (closeTrade) |
| 6 | G | Close Time | Data (closeTrade) |
| 7 | H | Trade Time | Formula |
| 8 | I | Strategy | Data (strategy) |
| 9 | J | Credit | Data (credit) |
| 10 | K | Debit | Data (debit) |
| 11 | L | Contracts | Data (contracts) |
| 12 | M | Net Profit | Formula |
| 13 | N | Open Fees | Data (open_fees) |
| 14 | O | Close Fees | Data (close_fees) |
| 15 | P | Profit | Formula |
| 16 | Q | Short Calls | Data (sold_call_strike) |
| 17 | R | Short Puts | Data (sold_put_strike) |
| 18 | S | Spread Width | Formula |
| 19 | T | Width | Data (width) |
| 20–27 | U–AB | Delta columns | Data (delta) |

This layout has NO ATR column and 2 extra columns at positions A–B, explaining both the +2 offset and the offset reduction from +2 to +1 after the ATR position.

### Discovery 3: ATR Field Missing from COLUMN_MAP

The JSON input contains an `ATR` field (e.g., `3.86`) which maps to CSV column O (ATR). The current code's `COLUMN_MAP` has no entry for ATR — the field is present in the input but silently discarded during writes.

**Evidence:**
* JSON input: `"ATR": 3.86` and `"ATR": 4.12`
* CSV column O header: "ATR" with values like `3.88`, `3.90`, `4.05`, `6.38`
* `COLUMN_MAP` in both `server.py:441` and `core_operations.py:637`: no "O" key for ATR (current "O" maps to `close_fees`)
* `values_dict` construction in `core_operations.py:802–893`: no ATR field extraction

### Discovery 4: Formula Columns (Must Not Overwrite)

Five columns contain Excel formulas that must be preserved (not overwritten by write operations):

| Column | Header | Formula Evidence | Derivation |
|:---:|---|---|---|
| B | Open Day | `MON`, `TUE`, etc. | `=TEXT(A{row}, "DDD")` — derived from Date of Purchase |
| F | Trade Time | `00:05:51`, `04:02:20` | `=E{row}-C{row}` or similar time difference |
| K | Net Profit | `$500.00` = `(0.20-0)×25×100` | `=(H{row}-I{row})*J{row}*100` |
| N | Profit | `$413.74` = `500.00-86.26` | `=K{row}-L{row}-M{row}` |
| R | Spread Width | `$125` (only for IC trades) | `=ABS(P{row}-Q{row})` when both strikes present |

**Verification for the 2/2/2026 VCCS trade:**
* Credit ($0.20) - Debit ($0) = $0.20 per contract
* $0.20 × 25 contracts × 100 multiplier = $500.00 → matches K column ✓
* $500.00 - $86.26 fees - $0 close fees = $413.74 → matches N column ✓

### Discovery 5: Corrected COLUMN_MAP for CSV Layout

Based on value-position matching for the CSV workbook layout:

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
    "O": "ATR",               # ATR — NEW FIELD (not in current code)
    "P": "sold_call_strike",  # Short Calls
    "Q": "sold_put_strike",   # Short Puts
    "S": "width",             # Width
}
```

**Changes from current code:**

| Field | Old Column | New Column | Change |
|---|:---:|:---:|---|
| `open_date` | C | A | Shifted left 2 |
| `open_time` | E | C | Shifted left 2 |
| `strategy` | I | G | Shifted left 2 |
| `credit` | J | H | Shifted left 2 |
| `debit` | K | I | Shifted left 2 |
| `contracts` | L | J | Shifted left 2 |
| `open_fees` | N | L | Shifted left 2 |
| `close_fees` | O | M | Shifted left 2 |
| `ATR` | (none) | O | **New field** |
| `sold_call_strike` | Q | P | Shifted left 1 |
| `sold_put_strike` | R | Q | Shifted left 1 |
| `width` | T | S | Shifted left 1 |

### Discovery 6: Field Transformations Applied Before Writing

The `values_dict` construction in `core_operations.py:802–893` applies these transformations:

| Field | Transformation | Example |
|---|---|---|
| `strategy` | `map_strategy_name()` → short code | `"Call Credit Spread"` → `"VCCS"` |
| `credit` | `credit_received / (contracts × 100)` if total | `375.0 / (25×100)` = `0.15` |
| `debit` | `debit_paid / (contracts × 100)` if total | Same logic |
| `sold_call_strike` | Extract from `sold_strikes[0]` if array | `[6995]` → `6995` |
| `sold_put_strike` | Extract from `sold_strikes[0]` for puts | Same logic |
| `width` | `abs(bought_strikes[0] - sold_strikes[0])` if missing | Calculated from strikes |
| `open_date` | Written as-is (string) | `"02/26/2026"` |
| `open_time` | Written as-is (string) | `"10:11:10"` |
| `ATR` | Not currently extracted | N/A — needs implementation |

**Note:** Dates and times are NOT converted to Excel serial numbers. They are written as raw strings and formatted by Excel's cell format.

## JSON Input Field Inventory

Complete field inventory from the sample JSON input structure:

```json
{
  "trade_results": [
    {
      "order_id": 442689429,
      "open_date": "02/26/2026",
      "open_time": "10:11:10",
      "strategy": "Call Credit Spread",
      "credit": 375.0,
      "debit": 0.0,
      "contracts": 25,
      "open_fees": 86.24,
      "close_fees": null,
      "sold_call_strike": 6995,
      "sold_put_strike": null,
      "bought_call_strike": 7010,
      "bought_put_strike": null,
      "sold_strikes": [6995],
      "bought_strikes": [7010],
      "calls_contracts": 25,
      "puts_contracts": 0,
      "width": 15,
      "ATR": 3.86,
      "net_profit": 288.76,
      "excel_row": 57
    }
  ]
}
```

**Field categorization:**

| Category | Fields | Written to Excel? |
|---|---|---|
| Written to COLUMN_MAP | `open_date`, `open_time`, `strategy`, `credit`, `debit`, `contracts`, `open_fees`, `close_fees`, `sold_call_strike`, `sold_put_strike`, `width` | Yes (11 fields) |
| Missing from COLUMN_MAP | `ATR` | **No** — should be added |
| Used for transformations | `sold_strikes`, `bought_strikes`, `calls_contracts`, `puts_contracts` | No — intermediate values |
| Metadata only | `order_id`, `excel_row`, `net_profit`, `bought_call_strike`, `bought_put_strike` | No — not stored in worksheet |

## Technical Scenarios

### Scenario A: Grouped Range PATCH Writes (Selected Approach)

Write contiguous column groups as single PATCH calls using `range(address='G{row}:J{row}')` with multi-cell value arrays.

**Contiguous groups for the corrected COLUMN_MAP (12 columns):**

| Group | Columns | Fields | Values Array |
|:---:|---|---|---|
| 1 | A | `open_date` | `[[date]]` |
| 2 | C | `open_time` | `[[time]]` |
| 3 | G:J | `strategy, credit, debit, contracts` | `[[strategy, credit, debit, contracts]]` |
| 4 | L:M | `open_fees, close_fees` | `[[open_fees, close_fees]]` |
| 5 | O:Q | `ATR, sold_call_strike, sold_put_strike` | `[[ATR, call_strike, put_strike]]` |
| 6 | S | `width` | `[[width]]` |

**Result:** 6 PATCH calls per trade (down from 11 per-cell calls). For 2 trades: 12 calls. For N trades: 6×N calls.

**Implementation example:**

```python
COLUMN_GROUPS = [
    ("A", ["open_date"]),
    ("C", ["open_time"]),
    ("G", ["strategy", "credit", "debit", "contracts"]),
    ("L", ["open_fees", "close_fees"]),
    ("O", ["ATR", "sold_call_strike", "sold_put_strike"]),
    ("S", ["width"]),
]

async def write_trade_row(client, workbook_url, sheet_name, headers, row, values_dict):
    """Write a single trade using grouped range PATCHes."""
    updated_cells = []
    errors = []

    for start_col, fields in COLUMN_GROUPS:
        values = [values_dict.get(f, "") for f in fields]

        # Skip group if all values are empty
        if all(v == "" or v is None for v in values):
            continue

        # Replace None with empty string for partial groups
        values = ["" if v is None else v for v in values]

        if len(fields) == 1:
            address = f"{start_col}{row}"
        else:
            end_col = chr(ord(start_col) + len(fields) - 1)
            address = f"{start_col}{row}:{end_col}{row}"

        url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"

        try:
            resp = await client.patch(
                url, headers=headers,
                json={"values": [values]},
                timeout=30.0
            )
            if resp.status_code in (200, 201):
                updated_cells.append(address)
            else:
                errors.append(f"{address}: {resp.text}")
        except Exception as e:
            errors.append(f"{address}: {str(e)}")

    return updated_cells, errors
```

**Advantages:**
* 45% fewer HTTP calls (6 vs 11 per trade)
* Contiguous writes are atomic per range — reduces partial-write risk
* Easier to reason about write ordering
* Natural fit for the column layout's contiguous data groups

**Limitations:**
* Empty fields within a group write empty strings to cells (cosmetic, not destructive)
* Group must be contiguous — non-contiguous columns require separate calls

### Scenario B: Per-Cell PATCH Writes (Current Approach)

Current implementation — one PATCH call per non-empty field in `COLUMN_MAP`.

**Result:** Up to 11 PATCH calls per trade (current code). With ATR added: up to 12 calls.

**Advantages:**
* Simple implementation — iterate dict entries
* Skips empty fields (no unnecessary writes)

**Limitations:**
* Highest HTTP call volume: 11–12 calls per trade
* Risk of Graph API throttling on large batches
* Each cell write is a separate network round-trip

### Scenario C: Graph API Batch Requests

Wrap all range PATCHes for a trade (or all trades) into a single `POST /$batch` request.

**Result:** 1 HTTP call per trade (containing 6 grouped PATCHes), or 1 HTTP call for all trades (up to 20 operations per batch).

**Advantages:**
* Minimum HTTP round-trips
* Single atomic request per trade/batch

**Limitations:**
* More complex implementation (construct JSON batch request body)
* Batch API has a 20-operation limit per request
* Error handling more complex (per-operation status codes within batch response)
* Graph API batch for Excel workbook operations has documented reliability issues

**Verdict:** Scenario A (Grouped Range PATCH) is the selected approach. It provides a meaningful reduction in API calls (45%) with straightforward implementation. Scenario C (Batch) is a potential future optimization but adds complexity that isn't justified until batch sizes exceed 3–4 trades.

## Dependent Mapping Updates

If the COLUMN_MAP is corrected for the CSV layout, these dependent mappings must also be updated:

### closeTrade Column Updates

| Current | Corrected | Field |
|:---:|:---:|---|
| F (close_date) | D (Close Day) | Date trade was closed |
| G (close_time) | E (Time OUT) | Time trade was closed |
| Search C (date lookup) | Search A (Date of Purchase) | Row identification |
| Search E (time lookup) | Search C (Time IN) | Row identification |

**Location:** `core_operations.py` — `close_trade_impl()` and `server.py` — `excel.closeTrade` tool

### updateTradeWithDelta Column Updates

| Current | Corrected | Field |
|:---:|:---:|---|
| U (9:30–11 Call) | T (C Delta at 10am) | Call delta at 10am checkpoint |
| V (9:30–11 Put) | U (P Delta at 10am) | Put delta at 10am checkpoint |
| W (11–12 Call) | V (C Delta at 11:30 am) | Call delta at 11:30am |
| X (11–12 Put) | W (P Delta at 11:30am) | Put delta at 11:30am |
| Y (1–2pm Call) | X (C Delta at 1:30pm) | Call delta at 1:30pm |
| Z (1–2pm Put) | Y (P Delta at 1:30pm) | Put delta at 1:30pm |
| AA (2:30–3:30 Call) | Z (C Delta at 3:00pm) | Call delta at 3pm |
| AB (2:30–3:30 Put) | AA (P Delta at 3:00pm) | Put delta at 3pm |

**Location:** `core_operations.py:27–36` — `DELTA_COLUMN_MAPPING` dict

### Row Search Column Updates

The `log_trades_impl()` function searches column C to find the last date row for appending new trades. If corrected to the CSV layout, it should search column A instead.

**Location:** `core_operations.py:760–800` — used range and column read logic

## Potential Next Research

* **Dual workbook support**: If both the current workbook (2 extra columns, no ATR) and the CSV workbook (starts at A, has ATR) are in active use, a configurable COLUMN_MAP is needed. Research whether both layouts are used simultaneously.
  * Reasoning: The code works for whatever workbook it currently targets. Changing the COLUMN_MAP to match the CSV would break the existing workbook.
  * Reference: Column offset analysis in Discovery 1
* **Graph API batch reliability**: Test whether `POST /$batch` with Excel workbook PATCH operations works reliably for production use.
  * Reasoning: Scenario C could further reduce API calls but needs validation
  * Reference: Microsoft Graph API batch documentation
* **Time format handling**: The JSON sends `10:11:10` (24h with seconds) but the CSV shows `10:11 AM` (12h without seconds). Investigate whether Excel cell formatting handles this conversion automatically.
  * Reasoning: Time rendering differences may require explicit formatting
  * Reference: CSV column C values vs JSON `open_time` values
* **ATR source validation**: ATR values differ slightly between CSV (3.90) and JSON (3.86) for the same trade date. Determine whether this represents a timing difference or different data sources.
  * Reasoning: If ATR updates after logging, the initial write value may differ from the CSV snapshot
  * Reference: CSV column O vs JSON ATR field

## Recommendations

### Selected Approach: Grouped Range PATCH with Corrected COLUMN_MAP

1. **Update COLUMN_MAP** to the corrected 12-column version (Discovery 5) — adding ATR at column O and shifting all columns to match the CSV layout.

2. **Implement grouped range writes** (Scenario A) — 6 PATCH calls per trade using contiguous column groups.

3. **Update dependent mappings** — closeTrade columns (F→D, G→E), delta columns (shift all left by 1), and row search column (C→A).

4. **Add ATR extraction** to `values_dict` construction — extract the `ATR` field from trade input and include it in the write payload.

5. **Consider a configurable COLUMN_MAP** — if both workbook layouts are in active use, make the column offsets configurable via environment variable or parameter.

### Risk Assessment

| Risk | Severity | Mitigation |
|---|---|---|
| Breaking existing workbook | High | Verify which workbook layout is currently in production before changing COLUMN_MAP |
| Formula column overwrites | High | Formula columns (B, F, K, N, R) are excluded from COLUMN_MAP — no risk |
| Empty fields in groups | Low | Empty strings written to cells are cosmetically identical to empty cells |
| Time format mismatch | Low | Test that Excel accepts `10:11:10` and formats it as `10:11 AM` via cell formatting |

### Implementation Impact

| File | Changes Required |
|---|---|
| [mcp-server/core_operations.py](mcp-server/core_operations.py) | Update COLUMN_MAP (line 637), update write loop (lines 893–916), add ATR to values_dict (line 802+), update close_trade_impl columns, update DELTA_COLUMN_MAPPING (line 27) |
| [mcp-server/server.py](mcp-server/server.py) | Update inline COLUMN_MAP (line 441), update inline write loop (lines 730–758). **Preferably**: refactor to delegate to `log_trades_impl()` (eliminate duplication) |
| [mcp-server/config.py](mcp-server/config.py) | No changes needed |

**Estimated scope:** ~100 lines modified across 2 files (or ~540 lines removed if inline logic in `server.py` is replaced with delegation).
