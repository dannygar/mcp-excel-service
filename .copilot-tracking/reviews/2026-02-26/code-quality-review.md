<!-- markdownlint-disable-file -->

# Code Quality Review — MCP Excel Service

| Field | Value |
|---|---|
| **Date** | 2026-02-26 |
| **Scope** | Full codebase quality and accuracy review |
| **Files Reviewed** | server.py, core_operations.py, excel_helpers.py, auth.py, graph_api.py, config.py, test_server.py, Dockerfile, pyproject.toml |
| **Related Plan** | N/A (general quality review) |
| **Changes Log** | N/A |
| **Research** | N/A |

---

## Summary

| Severity | Count |
|---|---|
| **Critical** | 2 |
| **Major** | 6 |
| **Minor** | 7 |
| **Total** | 15 |

**Overall Status:** ⚠️ Needs Rework

---

## Critical Findings

### C1: Massive code duplication — `excel_log_trades` MCP tool does not delegate to `log_trades_impl`

- **File:** `mcp-server/server.py` lines 383–815 (~430 lines)
- **Also in:** `mcp-server/core_operations.py` lines 640–965 (~325 lines)
- **Details:** The `excel_log_trades` MCP tool function contains a full inline implementation of the trade-logging logic (JSON parsing, trade sorting, field mapping, cell writing). The REST endpoint `api_log_trades` correctly delegates to `log_trades_impl()` in `core_operations.py`, which contains a nearly identical copy. This means:
  - Bug fixes applied to one copy may not be applied to the other.
  - The two implementations may silently diverge over time.
  - The other three MCP tools (`excel.updateRange`, `excel.closeTrade`, `excel.updateTradeWithDelta`) correctly delegate to their `*_impl` counterparts.
- **Fix:** Refactor `excel_log_trades` to parse the JSON string, then delegate to `log_trades_impl()` — matching the pattern used by the REST endpoint and the other MCP tools.

### C2: Test suite references nonexistent MCP tool `excel.updateRowByLookup`

- **File:** `mcp-server/test_server.py` — 11 references across lines 190, 208, 210, 232, 295, 297, 319, 360, 428, 435, 458
- **Details:** The test suite checks for and calls `excel.updateRowByLookup`, which does not exist in the current server. The current tools are: `excel.updateRange`, `excel.logTrades`, `excel.updateTradeWithDelta`, `excel.closeTrade`. This means:
  - `test_list_tools` will always fail (expects `excel.updateRowByLookup`).
  - `test_update_row_by_lookup_schema`, `test_update_row_with_date_lookup`, `test_mixed_value_types`, and `test_integration_update_row` all call a nonexistent tool.
  - The `test_log_trades_schema` test is missing required fields (`url`, `file_name`) so it would fail even with valid credentials.
- **Fix:** Update the test suite to use the actual tool names and correct request parameters.

---

## Major Findings

### M1: N+1 API calls for cell writes (performance)

- **Files:** `mcp-server/server.py` lines 752–780; `mcp-server/core_operations.py` lines 910–940
- **Details:** Each trade writes cells one at a time via individual PATCH requests. For a trade with 11 fields, that is 11 HTTP calls per trade. For 10 trades, that is 110 Graph API calls. The Graph API supports batch requests (up to 20 per batch) and range updates, which could reduce this to ~6 calls for 10 trades.
- **Impact:** Slow performance, potential throttling by Graph API (429 Too Many Requests).
- **Fix:** Use range-based writes (one PATCH per trade row) or Graph API batch endpoint `/$batch`.

### M2: Environment-loading code duplicated between `server.py` and `config.py`

- **Files:** `mcp-server/server.py` lines 43–66; `mcp-server/config.py` lines 14–29
- **Details:** Both files independently load `.env` files using the same `MCP_ENV`-based logic. This means `load_dotenv` is called twice with potentially different behavior depending on import order. Environment loading should happen in exactly one place.
- **Fix:** Consolidate environment loading into `config.py` and import from there in `server.py`.

### M3: Date-format ambiguity in `parse_date_string`

- **File:** `mcp-server/excel_helpers.py` lines 28–50
- **Details:** The function supports both `%m/%d/%Y` (US) and `%d-%m-%Y` (EU) formats. A date like `01-02-2025` will be parsed as January 2nd (matching `%m-%d-%Y` first at line 48), but the function also lists `%d-%m-%Y` which would interpret it as February 1st. The format `%m-%d-%Y` appears before `%d-%m-%Y` in the list, so the US format wins — but this is fragile and undocumented.
- **Fix:** Document the priority order or remove one of the ambiguous formats, or require the caller to specify the format.

### M4: `Pydantic` models defined but unused by MCP tool endpoints

- **File:** `mcp-server/server.py` lines 36, 818–870
- **Details:** `BaseModel` and `Field` are imported from Pydantic, and request model classes (`UpdateRangeRequest`, `LogTradesRequest`, `UpdateDeltaRequest`, `CloseTradeRequest`) are defined. However, the REST endpoint functions manually parse `await request.json()` and validate fields by hand — the Pydantic models are never instantiated or used for validation. This means typos in field names, wrong types, or extra fields are silently accepted.
- **Fix:** Either use the Pydantic models for request validation (e.g., `UpdateRangeRequest(**body)`) or remove them to reduce dead code.

### M5: Root `pyproject.toml` has wrong project metadata

- **File:** `pyproject.toml` (root)
- **Details:** The root `pyproject.toml` defines `name = "market-intel-mcp-server"` with description "Azure Functions MCP server for Market Intelligence APIs". This does not match the actual project (`mcp-excel-service`). It also lists `azure-functions` and `requests` as dependencies, which are not used by this project.
- **Fix:** Update to match the actual project or remove the root `pyproject.toml` if `mcp-server/pyproject.toml` is the canonical one.

### M6: Token cache uses global mutable state — not thread-safe

- **File:** `mcp-server/graph_api.py` lines 26–29, 88–92
- **Details:** `_token_cache` is a global dict mutated in `get_access_token()`. With multiple concurrent requests (uvicorn workers), this creates a race condition where two coroutines could both see an expired token and both request new tokens simultaneously. In an asyncio single-threaded model this is less severe, but if multi-worker mode is ever used it becomes a real bug.
- **Fix:** Use `asyncio.Lock` to guard token refresh, or use MSAL's built-in token cache.

---

## Minor Findings

### m1: PEP 8 lint warnings — imports not at top of file

- **File:** `mcp-server/server.py` lines 72–104
- **Details:** 9 lint errors for module-level imports below the `load_dotenv` block. This is intentional (env vars must be loaded before importing modules that read them at import time), but the linter flags it.
- **Fix:** Add `# noqa: E402` comments or restructure to defer env-dependent logic to functions.

### m2: `debugpy` listed as a runtime dependency

- **File:** `mcp-server/requirements.txt` line 12; `mcp-server/pyproject.toml` line 14
- **Details:** `debugpy` is installed in the Docker production image. It should be a dev-only dependency.
- **Fix:** Move to a `[dependency-groups] dev` section or a separate `requirements-dev.txt`.

### m3: Dockerfile HEALTHCHECK uses synchronous httpx

- **File:** `mcp-server/Dockerfile` lines 43–44
- **Details:** The healthcheck runs `python -c "import httpx; httpx.get(...)"`. This creates a new Python process for each health check (every 30s), which is heavyweight. A simpler curl or wget would be more efficient.
- **Fix:** Replace with `CMD curl -f http://localhost:3000/health || exit 1` (requires adding curl to the image) or use `python -c "import urllib.request; urllib.request.urlopen(...)"` from stdlib.

### m4: Unused variable `last_date_display` in `excel_log_trades`

- **File:** `mcp-server/server.py` line 587
- **Details:** `last_date_display = dt.strftime("%m/%d/%Y")` is assigned but only used in a logger message. The same pattern exists in `core_operations.py`. Not a bug, but static analyzers will flag it.

### m5: `excel_helpers.py` — `is_likely_date_string` not used internally

- **File:** `mcp-server/excel_helpers.py` lines 89–104
- **Details:** `is_likely_date_string` is called only from `compare_values_for_search` but is also exported. Not a bug, but if it's meant to be internal, prefix with underscore.

### m6: Sheet name and address not URL-encoded in Graph API calls

- **Files:** Multiple locations in `server.py` and `core_operations.py`
- **Details:** Sheet names and cell addresses are interpolated directly into URLs (e.g., `f".../{sheet_name}/range..."`). Sheet names with special characters (spaces, quotes, `#`) could produce invalid URLs. The Graph API expects these to be properly quoted.
- **Fix:** Use `urllib.parse.quote()` for sheet names in URL paths.

### m7: No input size limits on trades array

- **File:** `mcp-server/server.py` line 387; `mcp-server/core_operations.py` line 640
- **Details:** A caller could submit thousands of trades in a single call, creating thousands of Graph API requests and potentially hitting timeouts or rate limits. No maximum batch size is enforced.
- **Fix:** Add a configurable maximum (e.g., 50 trades per call) with a clear error message.

---

## Validation Results

| Check | Status |
|---|---|
| Python syntax (all .py files) | ✅ Pass |
| Pylint E402 (imports not at top) | ⚠️ 9 warnings in server.py (intentional) |
| Type checking | ⚠️ No mypy/pyright configured |
| Unit tests | ❌ Tests reference nonexistent tool (`excel.updateRowByLookup`) |
| Build (Docker) | ✅ Dockerfile is syntactically valid |

---

## Follow-Up Recommendations

### From review (discovered)

1. **Refactor `excel_log_trades`** to delegate to `log_trades_impl()` — eliminates ~400 lines of duplication and aligns with the pattern used by the other three tools.
2. **Update test_server.py** to test actual tools (`excel.logTrades`, `excel.updateRange`, `excel.updateTradeWithDelta`, `excel.closeTrade`).
3. **Batch cell writes** using range-based PATCH or Graph API `$batch` to reduce API calls from O(n*11) to O(n).
4. **Consolidate env loading** into `config.py` only.
5. **Fix root `pyproject.toml`** metadata to match the actual project.
6. **Use Pydantic models** for REST endpoint validation or remove them.
7. **URL-encode sheet names** in Graph API URLs.
8. **Add input size limits** for trade arrays.

### Deferred (out of scope)

- Add mypy/pyright configuration for static type checking.
- Add pytest-based unit tests with mocked Graph API responses.
- Consider MSAL library for token management instead of manual implementation.
- Add rate limiting/retry logic for Graph API calls (429 handling).
