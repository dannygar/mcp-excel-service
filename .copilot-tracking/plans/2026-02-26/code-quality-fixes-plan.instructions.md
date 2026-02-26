---
applyTo: '.copilot-tracking/changes/2026-02-26/code-quality-fixes-changes.md'
---
<!-- markdownlint-disable-file -->
# Implementation Plan: Code Quality Fixes — MCP Excel Service

## Overview

Resolve 15 code quality findings (2 critical, 6 major, 7 minor) identified in the 2026-02-26 review by refactoring MCP tool delegation, fixing tests, consolidating environment loading, optimizing Graph API calls, and cleaning up project configuration.

## Objectives

### User Requirements

* Fix all 15 findings from the code quality review — Source: `.copilot-tracking/reviews/2026-02-26/code-quality-review.md`

### Derived Objectives

* Establish consistent delegation pattern across all 4 MCP tools — Derived from: C1 shows `excel.logTrades` and `excel.updateRange` do not follow the delegation pattern used by the other 2 tools
* Ensure test suite can run without immediate failures — Derived from: C2 shows tests reference nonexistent tools
* Reduce Graph API call volume to prevent throttling — Derived from: M1 shows up to 11×N individual PATCH calls per logTrades invocation
* Remove dead code and incorrect metadata — Derived from: M4 (unused Pydantic models) and M5 (wrong root pyproject.toml)

## Context Summary

### Project Files

* `mcp-server/server.py` (1430 lines) — MCP tools, REST endpoints, Pydantic models, server startup
* `mcp-server/core_operations.py` (~965 lines) — `*_impl()` functions called by REST endpoints
* `mcp-server/excel_helpers.py` (~430 lines) — Date parsing, file resolution helpers
* `mcp-server/graph_api.py` (159 lines) — Token cache, Graph API headers, workbook URL builder
* `mcp-server/config.py` (222 lines) — Duplicate env loading, strategy name mapping
* `mcp-server/test_server.py` (641 lines) — Partially broken test suite
* `mcp-server/pyproject.toml` — Correct project metadata (includes debugpy as runtime dep)
* `mcp-server/requirements.txt` — Runtime dependencies (includes debugpy)
* `mcp-server/Dockerfile` — Multi-stage build with heavyweight healthcheck
* `pyproject.toml` (root) — Wrong project name and dependencies (copy-paste remnant)

### References

* `.copilot-tracking/reviews/2026-02-26/code-quality-review.md` — Source of all 15 findings
* `.copilot-tracking/research/2026-02-26/code-quality-fixes-research.md` — Codebase architecture analysis

### Standards References

* #file:../../.github/copilot-instructions.md — Project conventions (FastMCP decorators, async/await, JSON returns)

## Implementation Checklist

### [ ] Implementation Phase 1: Refactor MCP tool delegation (C1, M4)

<!-- parallelizable: false -->

* [ ] Step 1.1: Refactor `excel.logTrades` to delegate to `log_trades_impl()`
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 19-55)
* [ ] Step 1.2: Refactor `excel.updateRange` to delegate to `update_range_impl()`
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 57-91)
* [ ] Step 1.3: Use Pydantic models for REST endpoint validation or remove them (M4)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 93-121)

### [ ] Implementation Phase 2: Fix test suite (C2)

<!-- parallelizable: false -->

* [ ] Step 2.1: Update `test_list_tools` to expect actual tool names
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 125-149)
* [ ] Step 2.2: Replace `excel.updateRowByLookup` tests with tests for actual tools
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 151-198)
* [ ] Step 2.3: Fix `test_log_trades_schema` to include required `url` and `file_name` params
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 200-220)

### [ ] Implementation Phase 3: Optimize Graph API calls (M1)

<!-- parallelizable: false -->

* [ ] Step 3.1: Refactor `log_trades_impl()` to use row-range PATCH instead of per-cell PATCH
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 224-280)
* [ ] Step 3.2: Refactor `close_trade_impl()` to use single range write for columns F-G
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 282-308)

### [ ] Implementation Phase 4: Consolidate environment loading and config fixes (M2, M5)

<!-- parallelizable: true -->

* [ ] Step 4.1: Remove duplicate env-loading from `config.py`
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 312-340)
* [ ] Step 4.2: Fix root `pyproject.toml` metadata
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 342-368)

### [ ] Implementation Phase 5: Minor fixes and token cache safety (m1-m7, M6)

<!-- parallelizable: true -->

> **Note:** Line numbers in server.py will have shifted after Phase 1 removes ~440 lines. Apply edits relative to surrounding code patterns, not absolute line numbers.

* [ ] Step 5.1: Add `# noqa: E402` comments to imports after `load_dotenv` (m1)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 372-390)
* [ ] Step 5.2: Move `debugpy` to dev-only dependency (m2)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 392-414)
* [ ] Step 5.3: Replace Dockerfile healthcheck with stdlib `urllib` (m3)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 416-432)
* [ ] Step 5.4: Remove unused `last_date_display` variable or suppress warning (m4)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 434-446)
* [ ] Step 5.5: URL-encode sheet names in Graph API calls (m6)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 448-474)
* [ ] Step 5.6: Add input size limit for trades array (m7)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 476-500)
* [ ] Step 5.7: Add `asyncio.Lock` to token cache (M6)
  * Details: .copilot-tracking/details/2026-02-26/code-quality-fixes-details.md (Lines 502-536)

### [ ] Implementation Phase 6: Validation

<!-- parallelizable: false -->

* [ ] Step 6.1: Run full project validation
  * Verify Python syntax on all modified `.py` files (`python -m py_compile`)
  * Run test suite: `cd mcp-server && uv run python test_server.py --test health` (if server is running)
  * Verify Docker build: `docker build -t mcp-excel-test -f mcp-server/Dockerfile mcp-server/`
* [ ] Step 6.2: Fix minor validation issues
  * Iterate on any syntax errors or import issues
* [ ] Step 6.3: Report blocking issues
  * Document issues requiring additional research
  * Provide next steps for any blocking problems

## Planning Log

See [code-quality-fixes-log.md](.copilot-tracking/plans/logs/2026-02-26/code-quality-fixes-log.md) for discrepancy tracking, implementation paths considered, and suggested follow-on work.

## Dependencies

* Python 3.12+ with `uv` for dependency management
* FastMCP framework (server.py uses `@mcp.tool()` and `@mcp.custom_route()` decorators)
* `urllib.parse` (stdlib) — for URL-encoding sheet names
* `asyncio` (stdlib) — for Lock in token cache

## Success Criteria

* All 4 MCP tools delegate to `*_impl()` functions in `core_operations.py` — Traces to: C1
* Test suite references only existing tools and passes schema validation — Traces to: C2
* `log_trades_impl()` uses grouped range writes (≤6 PATCHes per trade, down from 11) — Traces to: M1
* Environment loading happens in exactly one place (`server.py`) — Traces to: M2
* Root `pyproject.toml` matches actual project — Traces to: M5
* All modified files pass `python -m py_compile` — Traces to: general quality
* Server starts successfully with `uv run python server.py` — Traces to: general quality
