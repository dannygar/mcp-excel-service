<!-- markdownlint-disable-file -->
# Implementation Details: Code Quality Fixes — MCP Excel Service

## Context Reference

Sources:
* `.copilot-tracking/reviews/2026-02-26/code-quality-review.md` (code quality review)
* `.copilot-tracking/research/2026-02-26/code-quality-fixes-research.md` (codebase research)

## Implementation Phase 1: Refactor MCP tool delegation (C1, M4)

<!-- parallelizable: false -->

### Step 1.1: Refactor `excel.logTrades` to delegate to `log_trades_impl()`

Replace the ~440-line inline implementation in `server.py` (lines 374-815) with a thin wrapper that:
1. Parses the `trades` JSON string into a Python list
2. Delegates to `log_trades_impl()` from `core_operations.py`
3. Wraps the result in `json.dumps()`
4. Handles `httpx.HTTPError` and generic `Exception`

Target pattern (matches `excel.updateTradeWithDelta` and `excel.closeTrade`):

```python
@mcp.tool(name="excel.logTrades")
async def excel_log_trades(
    url: str,
    file_name: str,
    sheet_name: str,
    trades: str
) -> str:
    """
    Log multiple trades to an Excel workbook.
    [keep existing docstring]
    """
    try:
        # Parse the trades JSON string into a list
        try:
            trades_list = json.loads(trades)
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

        # Delegate to core implementation
        result = await log_trades_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            trades=trades_list
        )
        return json.dumps(result, indent=2)

    except httpx.HTTPError as e:
        logger.error(f"HTTP error in logTrades: {e}")
        return json.dumps({
            "status": "error",
            "message": f"HTTP error: {str(e)}",
        }, indent=2)
    except Exception as e:
        logger.error(f"Error logging trades: {e}")
        return json.dumps({
            "status": "error",
            "message": str(e),
        }, indent=2)
```

Files:
* `mcp-server/server.py` — Replace lines 374-815 with ~35-line delegating wrapper (keep full docstring from lines 374-443)

Discrepancy references:
* Addresses C1 (critical code duplication)

Success criteria:
* `excel.logTrades` MCP tool produces identical JSON output to current implementation
* The inline `COLUMN_MAP`, `parse_trade_datetime()`, field extraction, and cell-write loops are removed from `server.py`
* `log_trades_impl()` in `core_operations.py` remains the single source of truth

Context references:
* `mcp-server/server.py` (Lines 374-815) — Current inline implementation to replace
* `mcp-server/core_operations.py` (Lines 621-965) — `log_trades_impl()` function that already exists
* `mcp-server/server.py` (Lines 920-960) — REST endpoint `api_log_trades` showing correct delegation pattern

Dependencies:
* None — `log_trades_impl` is already imported at server.py line 101

### Step 1.2: Refactor `excel.updateRange` to delegate to `update_range_impl()`

Replace the inline implementation in `server.py` (lines 268-371) with a thin wrapper matching the delegation pattern.

Target pattern:

```python
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
    [keep existing docstring]
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

        # Delegate to core implementation
        result = await update_range_impl(
            url=url,
            file_name=file_name,
            sheet_name=sheet_name,
            address=address,
            values=values_list
        )
        return json.dumps(result, indent=2)

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
```

Files:
* `mcp-server/server.py` — Replace lines 268-371 with ~35-line delegating wrapper

Discrepancy references:
* Addresses secondary part of C1 (updateRange also doesn't delegate)

Success criteria:
* `excel.updateRange` MCP tool produces identical JSON output
* Inline `resolve_excel_file_ids`, `build_workbook_url`, and PATCH logic removed from `server.py`

Context references:
* `mcp-server/server.py` (Lines 268-371) — Current inline implementation
* `mcp-server/core_operations.py` (Lines 541-618) — `update_range_impl()` function
* `mcp-server/server.py` (Lines 870-915) — REST endpoint `api_update_range` showing correct delegation

Dependencies:
* `update_range_impl` is already imported at server.py line 101

### Step 1.3: Remove unused Pydantic models (M4)

Remove the 4 Pydantic model classes that are defined but never used. Also remove the `pydantic` import since it will no longer be needed.

**Decision: Remove rather than adopt.** The REST endpoints have evolved with alternative field name handling (`trade_date` or `open_date` or `executed_date`) that Pydantic models don't capture without significant rework. Removing is cleaner.

Files:
* `mcp-server/server.py` — Remove lines 39 (`from pydantic import BaseModel, Field`), and lines ~823-870 (4 model class definitions)
* `mcp-server/requirements.txt` — Keep `pydantic` (FastMCP may depend on it internally)
* `mcp-server/pyproject.toml` — Keep `pydantic` in dependencies

Discrepancy references:
* Addresses M4 (Pydantic models defined but unused)

Success criteria:
* No `class UpdateRangeRequest`, `LogTradesRequest`, `UpdateDeltaRequest`, or `CloseTradeRequest` in server.py
* No `from pydantic import BaseModel, Field` in server.py
* Server starts without import errors (verify pydantic not needed elsewhere)

Context references:
* `mcp-server/server.py` (Lines 39, 823-870) — Import and model definitions to remove

Dependencies:
* Step 1.1 and 1.2 must complete first (they change line numbers)

## Implementation Phase 2: Fix test suite (C2)

<!-- parallelizable: false -->

### Step 2.1: Update `test_list_tools` to expect actual tool names

Change `expected_tools` list from `["excel.updateRowByLookup", "excel.updateRange"]` to the 4 actual tools.

```python
expected_tools = ["excel.updateRange", "excel.logTrades", "excel.updateTradeWithDelta", "excel.closeTrade"]
```

Files:
* `mcp-server/test_server.py` — Update `test_list_tools()` function (~line 165)

Success criteria:
* `test_list_tools` expects exactly the 4 tool names that exist in server.py

Context references:
* `mcp-server/test_server.py` (Lines 155-185) — Current `test_list_tools` function

Dependencies:
* None

### Step 2.2: Replace `excel.updateRowByLookup` tests with tests for actual tools

Replace the 4 test functions that call `excel.updateRowByLookup`:
* `test_update_row_by_lookup_schema` → `test_update_trade_with_delta_schema`
* `test_update_row_with_date_lookup` → `test_close_trade_schema`
* `test_mixed_value_types` → `test_log_trades_with_valid_params`
* `test_integration_update_row` → `test_integration_log_trades`

Each test should call the actual tool with correct parameters matching the tool's schema.

**`test_update_trade_with_delta_schema`** — calls `excel.updateTradeWithDelta` with:
```python
test_params = {
    "url": SYNTHETIC_DATA["sharepoint_url"],
    "file_name": SYNTHETIC_DATA["file_name"],
    "sheet_name": SYNTHETIC_DATA["sheet_name"],
    "trade_date": "01/06/2026",
    "trade_time": "10:30 AM",
    "sold_strike": "P 6855",
    "delta": 0.12,
    "delta_time": "10:45 AM",
}
```

**`test_close_trade_schema`** — calls `excel.closeTrade` with:
```python
test_params = {
    "url": SYNTHETIC_DATA["sharepoint_url"],
    "file_name": SYNTHETIC_DATA["file_name"],
    "sheet_name": SYNTHETIC_DATA["sheet_name"],
    "trade_date": "01/06/2026",
    "trade_time": "10:30 AM",
    "close_date": "01/06/2026",
    "close_time": "4:00 PM",
}
```

**`test_log_trades_with_valid_params`** — calls `excel.logTrades` with all required params:
```python
test_params = {
    "url": SYNTHETIC_DATA["sharepoint_url"],
    "file_name": SYNTHETIC_DATA["file_name"],
    "sheet_name": SYNTHETIC_DATA["sheet_name"],
    "trades": '[{"open_date": "01/06/2026", "open_time": "10:30 AM", "strategy": "VPCS", "credit": 0.25, "contracts": 25}]',
}
```

**`test_integration_log_trades`** — integration test using `excel.logTrades` with real credentials.

Files:
* `mcp-server/test_server.py` — Replace 4 test functions and the integration test function

Success criteria:
* No references to `excel.updateRowByLookup` remain in test_server.py
* All test functions call existing tools with correct parameter schemas

Context references:
* `mcp-server/test_server.py` (Lines 186-470) — Functions to replace
* `mcp-server/server.py` (Lines 112-197) — `excel.updateTradeWithDelta` parameter schema
* `mcp-server/server.py` (Lines 200-265) — `excel.closeTrade` parameter schema
* `mcp-server/server.py` (Lines 374-443) — `excel.logTrades` parameter schema

Dependencies:
* Step 2.1 (same file, sequential edits)

### Step 2.3: Fix `test_log_trades_schema` to include required params

The existing `test_log_trades_schema` (~line 500) sends only `trades`, `reference_date`, and `sheet_name`. It is missing required `url` and `file_name` params, and `reference_date` is not a valid parameter.

Update to:
```python
test_params = {
    "url": SYNTHETIC_DATA["sharepoint_url"],
    "file_name": SYNTHETIC_DATA["file_name"],
    "sheet_name": "December",
    "trades": '[{"open_date": "12/23/2025", "open_time": "10:30 AM", "strategy": "VPCS", "credit": 0.25, "contracts": 25}]',
}
```

Also update `run_all_tests()` to reference the renamed test functions.

Files:
* `mcp-server/test_server.py` — Update test_log_trades_schema params and run_all_tests list

Success criteria:
* `test_log_trades_schema` includes `url` and `file_name`
* `run_all_tests` references only valid test function names

Dependencies:
* Step 2.2 (renamed functions must be reflected in runner)

## Implementation Phase 3: Optimize Graph API calls (M1)

<!-- parallelizable: false -->

### Step 3.1: Refactor `log_trades_impl()` to use row-range PATCH

Replace the per-cell PATCH loop in `core_operations.py` (~lines 910-945) with a single range PATCH per trade row.

Current pattern (11 PATCHes per trade):
```python
for col_letter, field_name in COLUMN_MAP.items():
    cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{col_letter}{target_row}')"
    response = await write_client.patch(cell_url, headers=headers, json={"values": [[value]]})
```

New pattern (1 PATCH per trade):
```python
# Build a row array matching columns C through T
# COLUMN_MAP maps: C, E, I, J, K, L, N, O, Q, R, T
# Full range C:T spans columns C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T (18 columns)
# Place values in correct positions, use "" for unmapped columns

ALL_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]
row_values = []
for col in ALL_COLUMNS:
    field_name = COLUMN_MAP.get(col)
    if field_name:
        row_values.append(values_dict.get(field_name, ""))
    else:
        row_values.append("")  # Unmapped columns get empty string (preserves existing values? No — overwrite)

# IMPORTANT: Using empty string for unmapped columns will CLEAR existing data in D, F, G, H, M, P, S.
# Columns F, G = close_date, close_time (managed by closeTrade tool)
# Must preserve existing values by reading current row first OR by writing only mapped columns.

# Safer approach: Write only the mapped column ranges using multiple small range writes
# Group adjacent columns: C (alone), E (alone), I-L (4 cols), N-O (2 cols), Q-R (2 cols), T (alone)
# This reduces 11 calls to 6 calls per trade — a 45% reduction
```

**Selected approach: Grouped adjacent column ranges.** This preserves data in unmapped columns (D, F, G, H, M, P, S) without needing to read the current row first.

Column groups for range writes:
| Group | Columns | Fields | Range |
|---|---|---|---|
| 1 | C | open_date | `C{row}` |
| 2 | E | open_time | `E{row}` |
| 3 | I, J, K, L | strategy, credit, debit, contracts | `I{row}:L{row}` |
| 4 | N, O | open_fees, close_fees | `N{row}:O{row}` |
| 5 | Q, R | sold_call_strike, sold_put_strike | `Q{row}:R{row}` |
| 6 | T | width | `T{row}` |

This reduces API calls from 11 per trade to 6 per trade (45% reduction). For 10 trades: 60 calls instead of 110.

Files:
* `mcp-server/core_operations.py` — Refactor per-cell write loop in `log_trades_impl()` (~lines 910-945)

Discrepancy references:
* Partially addresses M1 (reduces calls by 45%, not 90%+)
* DD-01 in Planning Log: Full row-range write rejected to preserve unmapped column data

Success criteria:
* Each trade writes at most 6 PATCH requests instead of 11
* Data in columns D, F, G, H, M, P, S is not overwritten
* logTrades produces identical JSON responses

Context references:
* `mcp-server/core_operations.py` (Lines 900-945) — Current per-cell write loop
* `mcp-server/core_operations.py` (Lines 541-618) — `update_range_impl()` showing range PATCH pattern

Dependencies:
* Phase 1 must complete first (ensures MCP tool delegates to this function)

### Step 3.2: Refactor `close_trade_impl()` to use single range write

Replace the 2 individual PATCH calls for columns F and G with a single range PATCH for `F{row}:G{row}`.

Current:
```python
# PATCH F{row} with close_date
# PATCH G{row} with close_time
```

New:
```python
range_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='F{row}:G{row}')"
response = await client.patch(range_url, headers=headers, json={"values": [[close_date, close_time]]})
```

Files:
* `mcp-server/core_operations.py` — Refactor close_trade_impl() (~lines 478-530)

Success criteria:
* close_trade_impl writes 1 PATCH instead of 2
* Identical JSON response

Context references:
* `mcp-server/core_operations.py` (Lines 478-530) — Current 2-PATCH implementation

Dependencies:
* None (independent of Phase 1)

## Implementation Phase 4: Consolidate env loading and config fixes (M2, M5)

<!-- parallelizable: true -->

### Step 4.1: Remove duplicate env-loading from `config.py`

Remove the env-loading block from `config.py` (lines 7-28). `server.py` runs first and loads env vars before `config.py` is imported.

Replace:
```python
import os
import pathlib
from dotenv import load_dotenv

# Load environment variables from config folder
_project_root = pathlib.Path(__file__).parent.parent
_mcp_env = os.getenv("MCP_ENV", "local")
_env_file = _project_root / "config" / f".env.{_mcp_env}"
if _env_file.exists():
    load_dotenv(_env_file)
else:
    _env_local = _project_root / "config" / ".env.local"
    if _env_local.exists():
        load_dotenv(_env_local)
    else:
        load_dotenv()
```

With:
```python
# Environment variables are loaded by server.py before this module is imported.
# Do not call load_dotenv() here to avoid double-loading.
```

Also remove the `import pathlib` and `from dotenv import load_dotenv` if no longer needed.

Files:
* `mcp-server/config.py` — Remove lines 7-28 (imports `os`, `pathlib`, `dotenv` and env loading block). None of these are used by `map_strategy_name()` or `STRATEGY_MAPPING`.

Success criteria:
* `load_dotenv` is called exactly once (in server.py)
* `config.py` no longer imports `pathlib` or `dotenv`
* Server starts correctly with env vars loaded from server.py

Context references:
* `mcp-server/config.py` (Lines 7-28) — Duplicate env loading to remove
* `mcp-server/server.py` (Lines 43-66) — Primary env loading (kept)

Dependencies:
* None

### Step 4.2: Fix root `pyproject.toml` metadata

Replace the incorrect root `pyproject.toml` content with metadata matching the actual project. The canonical project config is `mcp-server/pyproject.toml`.

Replace:
```toml
[project]
name = "market-intel-mcp-server"
version = "0.1.0"
description = "Azure Functions MCP server for Market Intelligence APIs..."
dependencies = ["azure-functions", "requests", "python-dotenv"]
```

With:
```toml
[project]
name = "mcp-excel-service"
version = "1.0.0"
description = "Azure Container Apps-based MCP server providing Excel file manipulation capabilities for AI agents"
requires-python = ">=3.11"
dependencies = []
```

Remove the `[tool.hatch.build.targets.wheel]` section since the root is just a workspace marker.

Files:
* `pyproject.toml` (root) — Replace entire file content

Success criteria:
* Root pyproject.toml has correct project name and description
* No incorrect dependencies listed

Context references:
* `pyproject.toml` (root) (Lines 1-20) — Current wrong content
* `mcp-server/pyproject.toml` (Lines 1-5) — Canonical project config for reference

Dependencies:
* None

## Implementation Phase 5: Minor fixes (m1-m7)

<!-- parallelizable: true -->

### Step 5.1: Add `# noqa: E402` to imports after `load_dotenv` (m1)

Add `# noqa: E402` comments to the 9 import lines that follow the `load_dotenv` block in server.py (lines 72-101).

Files:
* `mcp-server/server.py` — Add `# noqa: E402` to lines 72-77 (httpx, FastMCP, starlette), 80-92 (auth, graph_api, excel_helpers, core_operations), 95 (config)

Success criteria:
* Linters no longer flag E402 on these imports

Dependencies:
* None (line numbers may shift after Phase 1; apply relative to surrounding code)

### Step 5.2: Move `debugpy` to dev-only dependency (m2)

Remove `debugpy>=1.8.0` from runtime dependencies in both files and add to dev group.

Files:
* `mcp-server/requirements.txt` — Remove `debugpy>=1.8.0` from main section, add `# Dev dependencies (not installed in production)` comment
* `mcp-server/pyproject.toml` — Move `debugpy>=1.8.0` from `dependencies` to `[dependency-groups] dev`

New `mcp-server/pyproject.toml` structure:
```toml
[project]
dependencies = [
    "fastmcp>=2.0.0",
    "httpx>=0.27.0",
    "python-dotenv>=1.0.0",
    "starlette>=0.40.0",
    "uvicorn>=0.30.0",
    "PyJWT>=2.8.0",
    "cryptography>=42.0.0",
]

[dependency-groups]
dev = [
    "debugpy>=1.8.0",
]
```

New `mcp-server/requirements.txt`:
```
# MCP Excel Server dependencies
fastmcp>=2.0.0
httpx>=0.27.0
python-dotenv>=1.0.0
starlette>=0.40.0
uvicorn>=0.30.0
pydantic>=2.0.0

# Authentication - Entra ID token validation
PyJWT>=2.8.0
cryptography>=42.0.0
```

Files:
* `mcp-server/pyproject.toml` — Move debugpy to dev group
* `mcp-server/requirements.txt` — Remove debugpy line

Success criteria:
* `debugpy` not installed in production Docker image
* `debugpy` still available for local development via `uv sync --group dev`

Dependencies:
* None

### Step 5.3: Replace Dockerfile healthcheck with stdlib `urllib` (m3)

Replace:
```dockerfile
CMD python -c "import httpx; httpx.get('http://localhost:3000/health')" || exit 1
```

With:
```dockerfile
CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:3000/health')" || exit 1
```

Files:
* `mcp-server/Dockerfile` — Update HEALTHCHECK CMD (lines 43-44)

Success criteria:
* Healthcheck works without importing httpx
* Docker build succeeds

Dependencies:
* None

### Step 5.4: Remove unused `last_date_display` variable (m4)

This variable exists only in the inline `excel_log_trades` which is being replaced in Step 1.1. If it also exists in `core_operations.py`, prefix the assignment with `_ = ` or remove the unused variable.

Files:
* `mcp-server/core_operations.py` — Check for `last_date_display` and either use it or remove it

Success criteria:
* No unused variable warnings for `last_date_display`

Dependencies:
* Step 1.1 (removes the server.py copy)

### Step 5.5: URL-encode sheet names in Graph API calls (m6)

Add `urllib.parse.quote(sheet_name, safe='')` wherever sheet names are interpolated into Graph API URLs.

Add import at top of affected files:
```python
from urllib.parse import quote as url_quote
```

Replace patterns like:
```python
f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"
```
With:
```python
f"{workbook_url}/worksheets/{url_quote(sheet_name, safe='')}/range(address='{address}')"
```

Files:
* `mcp-server/core_operations.py` — All Graph API URL constructions containing `sheet_name`
* `mcp-server/server.py` — Any remaining URL constructions after Phase 1 refactor (should be minimal since tools delegate)

Success criteria:
* Sheet names with spaces, quotes, or `#` produce valid Graph API URLs

Dependencies:
* Phase 1 (reduces the number of locations to update in server.py)

### Step 5.6: Add input size limit for trades array (m7)

Add a configurable maximum trade count at the top of `log_trades_impl()`:

```python
MAX_TRADES_PER_CALL = 50

async def log_trades_impl(url, file_name, sheet_name, trades: list) -> dict:
    if len(trades) > MAX_TRADES_PER_CALL:
        return {
            "status": "error",
            "message": f"Too many trades ({len(trades)}). Maximum is {MAX_TRADES_PER_CALL} per call.",
        }
    ...
```

Also add the same check in the MCP tool wrapper (after JSON parsing) for early rejection.

Files:
* `mcp-server/core_operations.py` — Add limit check at top of `log_trades_impl()`

Success criteria:
* Calls with >50 trades return a clear error without making any API calls

Dependencies:
* Step 1.1 (MCP tool delegates; limit in impl covers both MCP and REST paths)

### Step 5.7: Add `asyncio.Lock` to token cache (M6)

Add an `asyncio.Lock` to guard the token refresh in `get_access_token()`.

```python
import asyncio

_token_lock = asyncio.Lock()

async def get_access_token() -> str:
    global _token_cache
    
    # Check cache outside lock (fast path)
    if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 300:
        return _token_cache["access_token"]
    
    async with _token_lock:
        # Double-check inside lock (another coroutine may have refreshed)
        if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 300:
            return _token_cache["access_token"]
        
        # ... existing token acquisition logic ...
```

Files:
* `mcp-server/graph_api.py` — Add `import asyncio`, create `_token_lock`, wrap token refresh in lock

Success criteria:
* Only one coroutine can refresh the token at a time
* Existing behavior preserved for single-request scenarios

Context references:
* `mcp-server/graph_api.py` (Lines 26-123) — Token cache and get_access_token()

Dependencies:
* None

## Implementation Phase 6: Validation

<!-- parallelizable: false -->

### Step 6.1: Run full project validation

Execute all validation commands:
* `cd mcp-server && python -m py_compile server.py` — Verify syntax
* `cd mcp-server && python -m py_compile core_operations.py` — Verify syntax
* `cd mcp-server && python -m py_compile config.py` — Verify syntax
* `cd mcp-server && python -m py_compile graph_api.py` — Verify syntax
* `cd mcp-server && python -m py_compile test_server.py` — Verify syntax
* `cd mcp-server && python -m py_compile excel_helpers.py` — Verify syntax
* `docker build -t mcp-excel-test -f mcp-server/Dockerfile mcp-server/` — Verify Docker build (optional)

### Step 6.2: Fix minor validation issues

Iterate on syntax errors, import issues, and any compilation failures discovered in Step 6.1.

### Step 6.3: Report blocking issues

When validation failures require changes beyond minor fixes:
* Document the issues and affected files.
* Provide the user with next steps.
* Recommend additional research and planning rather than inline fixes.

## Dependencies

* Python 3.12+ with `uv`
* FastMCP framework
* `urllib.parse` (stdlib)
* `asyncio` (stdlib)

## Success Criteria

* All 4 MCP tools delegate to `*_impl()` counterparts
* No references to `excel.updateRowByLookup` in test_server.py
* `log_trades_impl()` uses grouped range writes (≤6 PATCHes per trade)
* `close_trade_impl()` uses single range write (1 PATCH)
* Environment loaded once in server.py only
* Root pyproject.toml has correct metadata
* All `.py` files pass `python -m py_compile`
