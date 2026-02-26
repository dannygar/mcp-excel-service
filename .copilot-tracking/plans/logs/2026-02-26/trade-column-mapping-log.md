<!-- markdownlint-disable-file -->
# Planning Log: Trade Column Mapping Fix

## Discrepancy Log

Gaps and differences identified between research findings and the implementation plan.

### Unaddressed Research Items

* DR-01: Dual workbook support — research recommends investigating whether both excel layouts (current 2-extra-column workbook and CSV layout) are in active use simultaneously
  * Source: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Potential Next Research, item 1)
  * Reason: Plan assumes the CSV layout is the target. If the existing workbook has a different layout, this change will break it. User must confirm which workbook is the production target before implementation.
  * Impact: High — incorrect assumption would write data to wrong columns in production

* DR-02: Graph API batch reliability — research suggests testing POST /$batch for Excel operations as a future optimization
  * Source: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Potential Next Research, item 2)
  * Reason: Excluded from current scope; grouped range writes (Scenario A) provide sufficient improvement. Batch can be evaluated as a follow-on.
  * Impact: Low — Scenario A is viable without batch API

* DR-03: Time format handling — research notes JSON sends "10:11:10" (24h) but CSV shows "10:11 AM" (12h)
  * Source: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Potential Next Research, item 3)
  * Reason: Assumed Excel cell formatting handles the conversion. Not validated. If times display incorrectly, an explicit format conversion would be needed.
  * Impact: Low — cosmetic issue, not data corruption

* DR-04: ATR source validation — research notes ATR values differ slightly between CSV and JSON
  * Source: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Potential Next Research, item 4)
  * Reason: Timing difference is most likely explanation. The plan writes the JSON ATR value as-is, which is correct for the moment the trade is logged.
  * Impact: Low — expected behavior for point-in-time data

### Plan Deviations from Research

* DD-01: Inline duplication elimination added as a derived objective
  * Research recommends: Updating COLUMN_MAP in both server.py and core_operations.py
  * Plan implements: Removing inline logic from server.py entirely, delegating to core_operations.log_trades_impl()
  * Rationale: Maintaining two copies of ~300 lines of identical logic is a maintenance risk. The delegation pattern is already established by closeTrade and updateRange. Fixing the duplication first (Phase 1) means COLUMN_MAP changes only need to be made once (Phase 2).

* DD-02: Phase ordering — refactor before column fix
  * Research recommends: Column fix as the primary action
  * Plan implements: Duplication elimination (Phase 1) before column correction (Phase 2)
  * Rationale: If COLUMN_MAP is fixed first in core_operations.py but server.py still has its own inline copy, both files would need identical changes. Eliminating the duplicate first ensures the column fix is applied in exactly one place.

## Implementation Paths Considered

### Selected: Grouped Range PATCH Writes with Refactor-First Approach

* Approach: Eliminate server.py duplication first, then fix COLUMN_MAP in core_operations.py, implement grouped range writes, and update dependent mappings
* Rationale: Single source of truth for column mapping; 45% fewer API calls; matches existing delegation patterns in the codebase
* Evidence: .copilot-tracking/research/2026-02-26/trade-column-mapping-research.md (Scenario A analysis, Dependent Mapping Updates)

### IP-01: In-Place Column Fix Without Refactoring

* Approach: Fix COLUMN_MAP in both server.py and core_operations.py without eliminating the duplication
* Trade-offs: Simpler change (just swap column letters), but maintains the dual-maintenance burden and requires identical changes in two files
* Rejection rationale: The duplication is already causing maintenance friction. With 12+ columns and grouped writes, the inline logic would diverge further. Fixing the root cause (duplication) first is worth the additional Phase 1 effort.

### IP-02: Graph API Batch Requests (Scenario C)

* Approach: Wrap all range PATCHes into POST /$batch calls for minimum HTTP round-trips
* Trade-offs: Maximum efficiency (1 HTTP call per trade), but more complex implementation, 20-operation batch limit, documented reliability issues with Excel batch operations
* Rejection rationale: Grouped range writes (Scenario A) provide 45% reduction with straightforward implementation. Batch API adds complexity without proportional benefit for typical 1–3 trade batch sizes. Can be added as follow-on work if performance becomes an issue.

### IP-03: Configurable COLUMN_MAP via Environment Variable

* Approach: Make column offsets configurable so both workbook layouts can be supported simultaneously
* Trade-offs: Supports dual-layout scenario, but adds configuration complexity and testing surface area
* Rejection rationale: No evidence that both layouts are used simultaneously. Adding configurability without a confirmed requirement adds unnecessary complexity. Documented as follow-on work (WI-01) pending user confirmation.

## Suggested Follow-On Work

Items identified during planning that fall outside current scope.

* WI-01: Configurable COLUMN_MAP for dual workbook support — If both Excel layouts (2-extra-column and CSV) are in production use, implement environment-variable-driven column offset configuration (High)
  * Source: Research Potential Next Research item 1; DR-01
  * Dependency: Confirm with user which workbook layout(s) are active

* WI-02: Graph API batch writes optimization — Evaluate POST /$batch reliability for Excel workbook operations and implement if viable for larger trade batches (Medium)
  * Source: Research Scenario C analysis; DR-02
  * Dependency: Phase 2 completion (grouped writes must be working first)

* WI-03: Time format validation — Test that Excel cell formatting correctly converts "10:11:10" (24h) to "10:11 AM" (12h) display format (Low)
  * Source: Research Potential Next Research item 3; DR-03
  * Dependency: Phase 5 integration testing

* WI-04: Unit test coverage for column mapping — Add automated tests for COLUMN_MAP, COLUMN_GROUPS, DELTA_COLUMN_MAPPING, and close_trade column references (Medium)
  * Source: Derived from plan — no existing test suite validates column references
  * Dependency: Phase 5 completion

## Validation Results

**Validation date:** 2026-02-26
**Validation status:** PASS_WITH_MINOR_FINDINGS

### Finding Summary

| # | Severity | Category | Description |
|---|---|---|---|
| VF-01 | Major | Completeness | `update_trade_with_delta_impl` search range not covered by any plan step |
| VF-02 | Minor | Accuracy | Step 1.2 describes work that is already done (REST endpoint already delegates) |
| VF-03 | Minor | Consistency | Line number references off by 1–12 lines in several plan/details locations |
| VF-04 | Minor | Completeness | MCP tool docstring column references in server.py not explicitly listed for update |
| VF-05 | Minor | Completeness | Impl function docstring column references in core_operations.py not mentioned |

### Detailed Findings

#### VF-01: Missing update for `update_trade_with_delta_impl` search range (Major)

The `update_trade_with_delta_impl()` function in [core_operations.py](mcp-server/core_operations.py#L221-L223) contains a hardcoded search range `C1:E{row_count}` (line 223) and column index comments referencing "Column C" (line 249) and "Column E" (line 250). This needs the same `A1:C{row_count}` update as `close_trade_impl()`.

**No plan step covers this update:**
* Step 3.1 covers only `DELTA_COLUMN_MAPPING` (the target write columns)
* Step 3.2 covers only `close_trade_impl` search range
* Step 3.3 covers only `log_trades_impl` date column search
* Step 4.2 checks only the MCP tool in server.py, not the impl function

**Impact:** After implementation, `update_trade_with_delta_impl` would still search columns C:E instead of A:C for row matching, causing delta updates to fail or match wrong rows.

**Recommendation:** Add a Step 3.4 (or extend Step 3.2) to update `update_trade_with_delta_impl` search range from `C1:E{row_count}` to `A1:C{row_count}` at core_operations.py lines 221–223, plus the column comments at lines 249–250.

#### VF-02: Step 1.2 describes already-completed work (Minor)

The plan's Step 1.2 says to "Refactor REST endpoint `api_log_trades` to call `log_trades_impl()` directly instead of calling through the MCP tool wrapper." However, the actual code at [server.py](mcp-server/server.py#L956) already calls `log_trades_impl()` directly:

```python
result = await log_trades_impl(
    url=body["url"],
    file_name=body["file_name"],
    sheet_name=body["sheet_name"],
    trades=trades_list
)
```

The research document itself notes: "REST endpoint `api_log_trades` at line 920 correctly delegates to `log_trades_impl()`."

**Impact:** Low — the implementer would simply confirm it's already done and move on.

**Recommendation:** Change Step 1.2 to a verification step ("Verify REST endpoint `api_log_trades` already delegates to `log_trades_impl()` — no changes needed").

#### VF-03: Line number references off by 1–12 lines (Minor)

Several line number references in the plan and details are slightly inaccurate:

| Reference | Plan/Details Says | Actual |
|---|---|---|
| COLUMN_MAP in server.py | L441–452 (plan) | L440–452 |
| COLUMN_MAP in core_operations.py | L638–649 (plan) | L637–649 |
| close_trade_impl search range | L414–419 (details) | L405–407 |
| server.py logTrades inline | L434–780 (details) | L440–790 |

**Impact:** Low — differences are small enough that the implementer can locate the correct code from surrounding context.

**Recommendation:** No action needed; line numbers are close enough for navigation.

#### VF-04: MCP tool docstring column references need explicit update list (Minor)

The server.py MCP tools contain hardcoded column references in docstrings that need updating:

* `excel.closeTrade` (lines 218–237): References columns C, E, F, G → should be A, C, D, E
* `excel.updateTradeWithDelta` (lines 128–164): References columns C, E, U–AB → should be A, C, T–AA
* `excel.logTrades` (lines 392–435): References "column C" → should be "column A"

Steps 4.1 and 4.2 implicitly cover this ("update them to match the corrected layout") but don't list the specific docstring lines.

**Impact:** Low — docstring-only changes, no runtime effect.

**Recommendation:** Add explicit docstring line references to Steps 4.1 and 4.2 details.

#### VF-05: Impl function docstring column references not mentioned (Minor)

The core_operations.py impl functions contain column references in their docstrings:

* `close_trade_impl` (lines 353–356): `(Column C)`, `(Column E)`, `(Column F)`, `(Column G)` → should be A, C, D, E
* `update_trade_with_delta_impl` (lines 135–136): `(Column C)`, `(Column E)` → should be A, C

**Impact:** Low — documentation accuracy only.

**Recommendation:** Include docstring updates as part of Steps 3.2 and the new Step 3.4.

### Coverage Validation

| Research Discovery | Plan Phase | Status |
|---|---|---|
| D1: Column offset mismatch | Phase 2, Step 2.1 | Covered ✓ |
| D2: Reconstructed layouts | Informs corrected COLUMN_MAP | Covered ✓ |
| D3: ATR field missing | Phase 2, Step 2.2 | Covered ✓ |
| D4: Formula columns | Excluded from all write paths | Covered ✓ |
| D5: Corrected COLUMN_MAP | Phase 2, Step 2.1 | Covered ✓ |
| D6: Field transformations | Maintained in values_dict | Covered ✓ |

| User Requirement | Plan Phase | Status |
|---|---|---|
| Fix COLUMN_MAP | Phase 2 | Covered ✓ |
| Add ATR field | Phase 2, Step 2.2 | Covered ✓ |
| Update dependent mappings | Phase 3 | Partially covered (see VF-01) |
| Document column mapping | Research + plan + details | Covered ✓ |
| Determine write strategy | Phase 2, Step 2.3 (Scenario A) | Covered ✓ |

### Dependency Validation

Phase ordering is correct:
* Phase 1 → Phase 2: Eliminates duplication first so COLUMN_MAP changes apply once ✓
* Phase 2 → Phase 3: COLUMN_MAP must be correct before dependent updates ✓
* Phase 3 steps marked parallelizable: Steps 3.1–3.3 are independent of each other ✓
* Phase 4 → depends on Phases 2–3: Verifies residual references after updates ✓
* Phase 5 → depends on all prior: Validation must come last ✓

### Risk Validation

* DR-01 (dual workbook risk): Documented as High impact, with WI-01 follow-on for configurable COLUMN_MAP ✓
* Formula column overwrites: Excluded from COLUMN_MAP — mitigated ✓
* Empty fields in grouped writes: Documented as Low — cosmetic only ✓

### Recommendations

1. **~~Add Step 3.4~~** — RESOLVED: Step 3.4 added to plan and details covering `update_trade_with_delta_impl` search range (`C1:E` → `A1:C`)
2. **~~Reclassify Step 1.2~~** — RESOLVED: Step 1.2 reclassified as verification-only in plan and details
3. **Add docstring update notes** to Steps 3.2, 4.1, and 4.2 for column references in docstrings and comments — Accepted as minor, implementer can address inline

### Resolution Summary

**Post-validation status:** PASS
* VF-01 (Major): Resolved — Step 3.4 added to Phase 3
* VF-02 (Minor): Resolved — Step 1.2 reclassified as verification
* VF-03 (Minor): Accepted — line numbers close enough for navigation
* VF-04 (Minor): Accepted — implementer will update docstrings alongside column changes
* VF-05 (Minor): Accepted — implementer will update docstrings alongside column changes
