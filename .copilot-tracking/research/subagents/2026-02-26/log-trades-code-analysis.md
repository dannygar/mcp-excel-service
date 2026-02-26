# Log Trades Code Analysis

**Research date:** 2026-02-26
**Status:** Complete
**Files analyzed:** `mcp-server/server.py`, `mcp-server/core_operations.py`, `mcp-server/config.py`, `mcp-server/excel_helpers.py`, `mcp-server/graph_api.py`

---

## 1. Parameter List for `log_trades_impl()`

**Location:** [core_operations.py](mcp-server/core_operations.py#L637-L650)

| Parameter | Type | Description |
|-----------|------|-------------|
| `url` | `str` | SharePoint/OneDrive URL to the document library |
| `file_name` | `str` | Excel file name with `.xlsx` extension |
| `sheet_name` | `str` | Worksheet name (e.g., "January") |
| `trades` | `list` | List of trade dictionaries (**already parsed** from JSON) |

### MCP Tool (`excel.logTrades`) Parameters

**Location:** [server.py](mcp-server/server.py#L382-L388)

| Parameter | Type | Description |
|-----------|------|-------------|
| `url` | `str` | SharePoint/OneDrive URL |
| `file_name` | `str` | Excel file name |
| `sheet_name` | `str` | Worksheet name |
| `trades` | `str` | **JSON string** (parsed inside the tool before processing) |

The MCP tool accepts `trades` as a **string** and calls `json.loads(trades)` to produce a list. The REST API endpoint and `log_trades_impl()` accept `trades` as an already-parsed list.

---

## 2. Column Mapping: JSON Field to Excel Cell

### COLUMN_MAP Definition

Both `server.py` (line ~444) and `core_operations.py` (line ~651) define an identical `COLUMN_MAP`:

| Excel Column | Field Key | JSON Input Field(s) | Description |
|:---:|---|---|---|
| **C** | `open_date` | `open_date`, `date`, `executed_date` | Date when trade was opened |
| **E** | `open_time` | `open_time`, `time`, `executed_time` | Time when trade was opened |
| **I** | `strategy` | `strategy` | Strategy short code (after mapping) |
| **J** | `credit` | `credit`, `credit_received` | Credit received per contract |
| **K** | `debit` | `debit`, `debit_paid` | Debit paid per contract |
| **L** | `contracts` | `contracts` | Number of contracts |
| **N** | `open_fees` | `open_fees`, `fees`, `total_fees` | Fees paid when opening |
| **O** | `close_fees` | `close_fees` | Fees paid when closing |
| **Q** | `sold_call_strike` | `sold_call_strike`, `sold_strikes[0]` (for call spreads) | Strike price for sold calls |
| **R** | `sold_put_strike` | `sold_put_strike`, `sold_strikes[0]` (for put spreads) | Strike price for sold puts |
| **T** | `width` | `width`, calculated as `abs(bought_strikes[0] - sold_strikes[0])` | Width between strikes |

### Columns Deliberately NOT Written by `logTrades`

| Excel Column | Field | Written By |
|:---:|---|---|
| **A** | (unknown) | Not written by any tool in this codebase |
| **B** | (unknown) | Not written by any tool in this codebase |
| **D** | (unknown — between open_date and open_time) | Not written by any tool in this codebase |
| **F** | Close date | `excel.closeTrade` → `close_trade_impl()` |
| **G** | Close time | `excel.closeTrade` → `close_trade_impl()` |
| **H** | (unknown — between close_time and strategy) | Not written by any tool in this codebase |
| **M** | (unknown — between contracts and open_fees) | Not written by any tool in this codebase |
| **P** | (unknown — between close_fees and sold_call_strike) | Not written by any tool in this codebase |
| **S** | (unknown — between sold_put_strike and width) | Not written by any tool in this codebase |
| **U-AB** | Delta values (8 columns for time-window × strike-type) | `excel.updateTradeWithDelta` → `update_trade_with_delta_impl()` |

---

## 3. Field Transformations

### 3.1. Strategy Name Mapping

**Location:** [config.py](mcp-server/config.py#L189-L222) — `map_strategy_name()`

The `strategy` field is passed through `map_strategy_name()` which:

1. **Short-code passthrough**: If the input is already a known short code (e.g., `"VPCS"`, `"IC"`), returns it as-is with correct casing.
2. **Exact match**: Lowercased input is looked up in `STRATEGY_MAPPING` dict (~50+ entries). Examples:
   - `"Iron Condor"` → `"IC"`
   - `"Put Credit Spread"` → `"VPCS"`
   - `"Covered Call"` → `"CC"`
3. **Keyword fuzzy match**: If no exact match, checks `STRATEGY_KEYWORDS` list. Each entry is a list of keywords that all must appear in the input. First matching pattern wins.
4. **No match**: Returns original string unchanged.

### 3.2. Credit/Debit Per-Contract Conversion

When `credit` is missing but `credit_received` is provided:

```python
credit = credit_received / (contracts * 100)
```

Same logic for `debit` when `debit_paid` is provided:

```python
debit = debit_paid / (contracts * 100)
```

This converts total dollar amounts to per-contract option prices (dividing by 100 for options multiplier and by contract count).

### 3.3. Sold Strike Extraction from Array

When `sold_call_strike` or `sold_put_strike` is missing but `sold_strikes` array is present:

- **Call strikes**: Extracted from `sold_strikes[0]` if the strategy contains `"call"` OR `calls_contracts > 0` with `puts_contracts == 0`.
- **Put strikes**: Extracted from `sold_strikes[0]` if the strategy contains `"put"` OR `puts_contracts > 0` with `calls_contracts == 0`.

### 3.4. Width Calculation

When `width` is missing but both `sold_strikes` and `bought_strikes` arrays are present:

```python
width = abs(bought_strikes[0] - sold_strikes[0])
```

### 3.5. Alternative Field Name Resolution

| Canonical Field | Accepted Alternatives |
|---|---|
| `open_date` | `date`, `executed_date` |
| `open_time` | `time`, `executed_time` |
| `credit` | `credit_received` (with per-contract conversion) |
| `debit` | `debit_paid` (with per-contract conversion) |
| `open_fees` | `fees`, `total_fees` |
| `sold_call_strike` | `sold_strikes[0]` (for call strategies) |
| `sold_put_strike` | `sold_strikes[0]` (for put strategies) |
| `width` | Calculated from `sold_strikes[0]` and `bought_strikes[0]` |

### 3.6. No Date/Time Formatting

Dates and times are written to Excel **as-is** (raw string values). No conversion to Excel serial numbers or reformatting occurs during writes. The `parse_date_string()` and `excel_serial_to_date()` helpers are used only for **reading** dates when searching for the last occupied row.

---

## 4. Exact Sequence of Graph API Calls Per Trade Batch

For a batch of N trades, the following Graph API calls are made:

### Phase 1: Setup (once per batch)

1. **GET** `token endpoint` — Acquire/refresh OAuth token (via `get_graph_headers()`)
2. **GET** `{GRAPH_API_BASE}/sites/{hostname}:{site_path}` — Resolve site ID
3. **GET** `{GRAPH_API_BASE}/sites/{site_id}/drives` — List drives to find document library
4. **GET** `{GRAPH_API_BASE}/drives/{drive_id}/root:/{file_name}` — Resolve file item ID
5. **GET** `{workbook_url}/worksheets/{sheet_name}/usedRange` — Get used range (row count)
6. **GET** `{workbook_url}/worksheets/{sheet_name}/range(address='C1:C{rowCount}')` — Read column C to find last date row

### Phase 2: Per-Trade Writes (repeated for each trade)

For each trade `i` (0-indexed), the target row is `last_date_row + 1 + i`. Then for each non-empty field in `COLUMN_MAP`:

7. **PATCH** `{workbook_url}/worksheets/{sheet_name}/range(address='{col}{target_row}')` — Write single cell

   Body: `{"values": [[value]]}`

Each cell is written individually with a separate PATCH call. For a trade with all 11 fields populated, that is **up to 11 PATCH calls per trade**. Empty/null fields are skipped.

### Write Order

Cells are written in the iteration order of the `COLUMN_MAP` dictionary:

```
C → E → I → J → K → L → N → O → Q → R → T
```

(Python 3.7+ dicts preserve insertion order.)

---

## 5. Sorting Behavior

Before any writes, the entire `trades` list is sorted by:

1. `open_date` (ascending, parsed via `parse_date_string()`)
2. `open_time` (ascending, parsed via time format parsing)

Trades with unparseable dates/times sort to the end (`datetime.max`).

---

## 6. Error Handling

| Scenario | Behavior |
|---|---|
| `trades` is not a list | Returns `{"status": "error"}` immediately |
| Empty trades list | Returns `{"status": "warning", "message": "No trades provided"}` |
| URL resolution fails | Returns `{"status": "error"}` with resolution error message |
| Used range request fails | Returns `{"status": "error"}` with HTTP error |
| Column C read fails | Returns `{"status": "error"}` with HTTP error |
| No valid date found in column C | Falls back to searching for "Date of Purchase" header; if not found, returns error |
| Individual cell PATCH fails | Error recorded in `cell_errors` list; trade marked as failed but processing continues |
| HTTP exception on cell write | Caught by generic `except Exception`, added to `cell_errors` |
| Some trades fail, some succeed | Returns `{"status": "partial_success"}` with both `results` and `errors` arrays |
| All trades fail | Returns `{"status": "error"}` |
| Missing/null field value | **Silently skipped** — the cell is simply not written (no PATCH call made) |
| Unknown strategy name | Written as-is (no mapping applied, original string used) |

---

## 7. Discrepancies Between `server.py` and `core_operations.py`

### Finding: The MCP tool has **completely duplicated inline logic**

The `excel.logTrades` MCP tool in `server.py` (lines 382-815) does **NOT** delegate to `log_trades_impl()`. Instead, it contains a full copy of the same logic inline. The code is functionally identical.

The REST API endpoint (`api_log_trades` at line 920) **does** delegate to `log_trades_impl()` correctly.

### Specific Duplication Points

| Aspect | `server.py` (MCP tool) | `core_operations.py` (`log_trades_impl`) |
|---|---|---|
| COLUMN_MAP | Defined inline at ~line 444 | Defined inline at ~line 651 |
| `parse_trade_datetime()` | Defined as nested function at ~line 478 | Defined as nested function at ~line 683 |
| Sort logic | Inline | Inline |
| File resolution | Inline | Inline |
| Used range + column C read | Inline | Inline |
| Last-date-row search | Inline | Inline |
| Field transformation logic | Inline (lines ~618-720) | Inline (lines ~802-893) |
| Cell write loop | Inline | Inline |
| Response construction | Returns `json.dumps(dict)` | Returns `dict` directly |

### Key Difference

- `server.py` MCP tool: Accepts `trades` as a **string**, parses JSON, returns **JSON string**.
- `core_operations.py`: Accepts `trades` as a **list**, returns **dict** (no JSON serialization).

### Risk

Any bug fix or feature change must be applied to **both** locations, or they will diverge. This is a significant code quality concern.

---

## 8. Delta Column Mapping (Reference — `updateTradeWithDelta`)

For completeness, the delta columns written by `update_trade_with_delta_impl()`:

| Time Window | Call Column | Put Column |
|---|:---:|:---:|
| 9:30 AM – 11:00 AM | **U** | **V** |
| 11:00 AM – 12:00 PM | **W** | **X** |
| 1:00 PM – 2:00 PM | **Y** | **Z** |
| 2:30 PM – 3:30 PM | **AA** | **AB** |

---

## 9. Close Trade Columns (Reference — `closeTrade`)

Written by `close_trade_impl()`:

| Excel Column | Field |
|:---:|---|
| **F** | Close date |
| **G** | Close time |

---

## Recommended Next Research

1. **Obtain the actual Excel template** to verify columns A, B, D, H, M, P, S — these are never written by any tool but may contain formulas, headers, or manually-entered data.
2. **Investigate the duplication** between `server.py` inline logic and `core_operations.py` — determine if the MCP tool should be refactored to delegate to `log_trades_impl()` (as the REST endpoint already does).
3. **Test edge cases**: What happens when `credit_received` division produces floating point precision issues? The per-contract conversion `credit_received / (contracts * 100)` could produce values like `0.30000000000000004`.
4. **Rate limiting**: Each trade makes up to 11 sequential PATCH calls. For large batches, this could hit Graph API throttling. Consider batch/range writes.
