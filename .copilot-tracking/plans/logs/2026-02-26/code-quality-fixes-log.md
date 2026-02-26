<!-- markdownlint-disable-file -->
# Planning Log: Code Quality Fixes — MCP Excel Service

## Discrepancy Log

Gaps and differences identified between research findings and the implementation plan.

### Unaddressed Research Items

* DR-01: Date format ambiguity in `parse_date_string` (M3)
  * Source: `.copilot-tracking/reviews/2026-02-26/code-quality-review.md` — M3
  * Reason: Changing date format priority could break existing Excel files. Requires data audit to determine which formats are actually used before changing behavior.
  * Impact: low — Current US-first priority works for the project's known use case.

* DR-02: `is_likely_date_string` internal naming convention (m5)
  * Source: `.copilot-tracking/reviews/2026-02-26/code-quality-review.md` — m5
  * Reason: Renaming a function that may be imported externally is a breaking change without a clear benefit. Deferred to avoid scope creep.
  * Impact: low — Cosmetic naming convention only.

### Plan Deviations from Research

* DD-00: Review states `excel.updateRange` delegates — it does not
  * Source: Code quality review C1 details claim "The other three MCP tools correctly delegate"
  * Actual: `excel.updateRange` (server.py ~line 279) inlines resolve/build/PATCH logic
  * Plan addresses: Step 1.2 refactors `excel.updateRange` to delegate to `update_range_impl()`
  * Impact: medium — Without Step 1.2, one tool would remain inconsistent

* DD-01: Graph API optimization uses grouped ranges instead of full-row write
  * Research recommends: Single full-row PATCH for `C{row}:T{row}` per trade (1 call/trade)
  * Plan implements: 6 grouped range PATCHes per trade (C, E, I:L, N:O, Q:R, T)
  * Rationale: A full-row write would overwrite columns D, F, G, H, M, P, S with empty strings, destroying data managed by other tools (e.g., closeTrade writes F and G). The grouped approach preserves existing data without requiring a read-before-write pattern.

* DD-02: Pydantic models removed instead of adopted
  * Research recommends: Use Pydantic models for REST endpoint validation
  * Plan implements: Remove Pydantic models entirely
  * Rationale: REST endpoints have evolved with complex alternative field name handling (e.g., `trade_date` or `open_date` or `executed_date`) that would require significant Pydantic model rework (custom validators, field aliases). The manual validation pattern is adequate and matches FastMCP's internal approach. Pydantic dependency is kept since FastMCP may use it internally.

## Implementation Paths Considered

### Selected: Incremental refactor with delegation pattern

* Approach: Refactor MCP tools to delegate to existing `*_impl()` functions in core_operations.py, fix tests to reference actual tools, optimize writes with grouped ranges, and clean up config.
* Rationale: Lowest risk — preserves all existing behavior, addresses the highest-severity findings first, and each phase can be validated independently.
* Evidence: `.copilot-tracking/research/2026-02-26/code-quality-fixes-research.md` (Section: MCP Tool Delegation Pattern) — 2 of 4 tools already follow this pattern successfully.

### IP-01: Major architectural restructure

* Approach: Restructure into router/handler/service layers, introduce proper dependency injection, add pytest with mocked Graph API, adopt Pydantic for all data models.
* Trade-offs: More maintainable long-term, but breaks all existing code, requires significant testing infrastructure setup, and high risk of introducing new bugs.
* Rejection rationale: Scope far exceeds the 15 review findings. Should be a separate planning effort if desired.

### IP-02: Graph API `$batch` endpoint for maximum efficiency

* Approach: Use Graph API batch requests (`/$batch`) to combine up to 20 operations per HTTP call, reducing logTrades from 11×N to ceil(11×N/20) calls.
* Trade-offs: Maximum API call reduction (82%+ for 10 trades), but requires implementing JSON batch payload construction, response parsing, partial failure handling, and retry logic for individual batch items.
* Rejection rationale: Significantly more complex than grouped range writes. The grouped approach achieves 45% reduction with minimal code change. Batch endpoint is recommended as follow-on work (WI-01).

## Suggested Follow-On Work

Items identified during planning that fall outside current scope.

* WI-01: Implement Graph API `$batch` endpoint for logTrades — Reduces per-trade API calls from 6 to ~1 by batching multiple range PATCHes into single HTTP calls (high priority)
  * Source: M1 review finding, IP-02 evaluation
  * Dependency: Phase 3 completion (grouped ranges as baseline)

* WI-02: Add pytest framework with mocked Graph API — Replace current `test_server.py` with proper pytest suite using `httpx.MockTransport` or `respx` for Graph API mocking (medium priority)
  * Source: C2 review finding, general test quality
  * Dependency: Phase 2 completion (tests reference correct tools)

* WI-03: Add mypy/pyright static type checking — Configure type checker in CI to catch type errors early (low priority)
  * Source: Validation Results in code quality review
  * Dependency: None

* WI-04: Implement MSAL for token management — Replace manual token cache in graph_api.py with Microsoft Authentication Library for automatic token management and rotation (low priority)
  * Source: M6 review finding, deferred recommendations
  * Dependency: None

* WI-05: Add retry logic for Graph API 429 responses — Implement exponential backoff when Graph API returns "Too Many Requests" (medium priority)
  * Source: M1 review finding context
  * Dependency: Phase 3 completion
