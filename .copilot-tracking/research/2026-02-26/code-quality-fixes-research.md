<!-- markdownlint-disable-file -->
# Research: Code Quality Fixes — MCP Excel Service

## Source

- Code quality review: `.copilot-tracking/reviews/2026-02-26/code-quality-review.md`
- Codebase analysis via subagent research (2026-02-26)

## Codebase Architecture

### Module Map

| Module | Lines | Responsibility |
|---|---|---|
| `mcp-server/server.py` | ~1430 | MCP tools, REST endpoints, Pydantic models, health check, server startup |
| `mcp-server/core_operations.py` | ~965 | `*_impl()` functions called by REST endpoints |
| `mcp-server/excel_helpers.py` | ~430 | Date parsing, URL parsing, file resolution |
| `mcp-server/graph_api.py` | ~165 | Token cache, headers, workbook URL builder |
| `mcp-server/auth.py` | ~358 | Entra ID JWT validation, middleware |
| `mcp-server/config.py` | ~222 | Env loading (duplicate), strategy mapping |
| `mcp-server/test_server.py` | ~641 | MCP tool tests (partially broken) |

### MCP Tool Delegation Pattern

Three tools (`excel.updateTradeWithDelta`, `excel.closeTrade`) follow this delegation pattern:

```python
@mcp.tool(name="excel.TOOLNAME")
async def excel_tool_name(params...) -> str:
    try:
        result = await TOOL_impl(params...)  # Delegates to core_operations
        return json.dumps(result, indent=2)
    except httpx.HTTPError as e:
        return json.dumps({"status": "error", "message": f"HTTP error: {str(e)}"})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})
```

**Non-conforming tools:**
- `excel.updateRange` (server.py lines 268-371): Inlines logic instead of calling `update_range_impl()`
- `excel.logTrades` (server.py lines 374-815): ~440 lines inline; does NOT call `log_trades_impl()`

### REST Endpoint Delegation Pattern

All 4 REST endpoints correctly delegate to `*_impl()` functions in `core_operations.py`:
- `api_update_range` → `update_range_impl()`
- `api_log_trades` → `log_trades_impl()`
- `api_update_trade_with_delta` → `update_trade_with_delta_impl()`
- `api_close_trade` → `close_trade_impl()`

## Findings by Review Item

### C1: `excel_log_trades` duplication (Critical)

- **MCP tool**: server.py lines 374-815 (~440 lines inline)
- **core_operations.py**: `log_trades_impl()` lines 621-965 (~345 lines)
- The MCP tool parses a JSON string for `trades`, then performs all logic inline
- The REST endpoint correctly calls `log_trades_impl()` which takes `trades: list`
- **Fix pattern**: Parse JSON string → list, then call `log_trades_impl(url, file_name, sheet_name, trades_list)`
- **Also fix**: `excel.updateRange` (server.py lines 268-371) doesn't delegate to `update_range_impl()` either

### C2: Test suite references nonexistent tool (Critical)

- `excel.updateRowByLookup` does not exist — replaced/removed at some point
- Affected tests in test_server.py:
  - `test_list_tools()` ~line 165: expects `excel.updateRowByLookup` in tool list
  - `test_update_row_by_lookup_schema()` ~line 186
  - `test_update_row_with_date_lookup()` ~line 320
  - `test_mixed_value_types()` ~line 367
  - `test_integration_update_row()` ~line 459
- `test_log_trades_schema()` ~line 500: missing `url` and `file_name` required fields
- No tests exist for: `excel.updateTradeWithDelta`, `excel.closeTrade`
- Current actual tools: `excel.updateRange`, `excel.logTrades`, `excel.updateTradeWithDelta`, `excel.closeTrade`

### M1: N+1 API calls for cell writes (Major)

- **logTrades**: Uses `COLUMN_MAP` with up to 11 columns → 11 individual PATCH per trade
- **closeTrade**: 2 individual PATCHes (columns F and G)
- **Existing efficient pattern**: `update_range_impl()` uses single PATCH with 2D array for range update
- **Graph API supports**: Single range PATCH `worksheets/{sheet}/range(address='A{row}:K{row}')` with `{"values": [[v1, v2, ...]]}`
- Would reduce 11×N calls to N calls for logTrades, 2→1 for closeTrade

### M2: Duplicate env loading (Major)

- `server.py` lines 43-66: Reads MCP_ENV, constructs dotenv path, calls `load_dotenv()`
- `config.py` lines 11-25: Identical logic
- Both run at import time; `server.py` runs first, `config.py` runs when imported

### M3: Date format ambiguity (Major)

- `excel_helpers.py` lines 28-50: `parse_date_string()` tries 5 formats in order
- `%m-%d-%Y` appears before `%d-%m-%Y` → US format wins for ambiguous dates
- No documentation of priority; fragile for non-US users

### M4: Pydantic models unused (Major)

- 4 models defined at server.py lines ~818-879: `UpdateRangeRequest`, `LogTradesRequest`, `UpdateDeltaRequest`, `CloseTradeRequest`
- REST endpoints manually parse `request.json()` and validate with loops
- Models are never instantiated

### M5: Root pyproject.toml wrong metadata (Major)

- Root `pyproject.toml`: `name = "market-intel-mcp-server"`, description references "Azure Functions", dependencies include `azure-functions` and `requests`
- Canonical config is `mcp-server/pyproject.toml`

### M6: Token cache not thread-safe (Major)

- `graph_api.py` lines 28-31: Global `_token_cache` dict
- No `asyncio.Lock` guard on token refresh
- Low risk in single-worker asyncio, higher risk if multi-worker

### m1-m7: Minor findings

- **m1**: 9 E402 lint warnings in server.py (intentional — imports after dotenv load)
- **m2**: `debugpy` in runtime deps (requirements.txt line 12, pyproject.toml line 14)
- **m3**: Dockerfile healthcheck uses `python -c "import httpx; httpx.get(...)"` (heavyweight)
- **m4**: Unused variable `last_date_display` in server.py line 587
- **m5**: `is_likely_date_string` could be prefixed with `_` if internal
- **m6**: Sheet names not URL-encoded in Graph API URLs
- **m7**: No input size limits on trades array

## Key Implementation Dependencies

- `log_trades_impl()` takes `trades: list` (already parsed) — MCP tool must parse JSON string first
- `update_range_impl()` takes `values: list` — MCP tool must parse JSON string first
- Both `*_impl` functions handle `resolve_excel_file_ids()` internally
- REST endpoints handle alternative field name mappings before calling `*_impl`
- `COLUMN_MAP` is defined in both server.py (~line 446) and core_operations.py (~line 640)
- Strategy mapping logic is in `config.py` (`map_strategy_name()`)

## Constraints

- No formal test infrastructure (no pytest config, no CI)
- Python 3.12 with `uv` for dependency management
- FastMCP framework for MCP tools and custom routes
- Graph API for all Excel operations
- Entra ID authentication (optional, controlled by env var)
