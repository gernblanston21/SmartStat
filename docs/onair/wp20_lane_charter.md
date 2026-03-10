# WP-20 Runtime-Bridge Lane Charter (Target-02)

Status date: `2026-03-10`  
Scope: `WP-20 Target-02 governance gate (docs/tests only)`  
Implementation state: `NOT STARTED`

## Purpose

Define the concrete runtime-bridge lane charter that any future WP-20
implementation branch must follow before runtime-bridge code begins.

This charter does not authorize implementation.

## Branch Isolation Expectations

1. Runtime-bridge work must execute on a dedicated WP-20 branch/lane.
2. `feature/semantic-layer` remains a governance/docs/tests lane for gating.
3. Runtime-bridge branch ownership, merge ownership, and rollback ownership must
   be declared before first code change.
4. Cross-lane cherry-picks must be explicitly reviewed and documented.
5. No hidden bridge work may be mixed into unrelated lanes.

## Allowed Upstream Inputs (Read-Only)

1. WP-17 captured-plan artifacts and contract/schema evidence.
2. WP-18 validation outputs:
   - `validation_result`
   - `rule_evaluations`
   - refusal diagnostics
   - deterministic identities
   - `semantic_interpretation`
3. WP-19 projection and adapter contract surfaces.

## Allowed Implementation Surface Categories (Future WP-20 Lane)

Only the categories below are pre-approved for a future implementation lane:

1. `Bridge Contract Surface`
- Explicit runtime-bridge contract modules/interfaces.
- Typed boundary definitions between validated artifacts and bridge execution.

2. `Bridge Orchestration Surface`
- Controlled orchestration code for bridge flow sequencing.
- Explicit guard points for fail-closed behavior.

3. `Safety and Guardrail Surface`
- Fail-closed checks, boundary assertions, and explicit refusal handling.
- Deterministic diagnostics needed for safety/auditability.

4. `Evidence Harness Surface`
- WP-20 regression/evidence harnesses and artifact emitters under `tests/wp-20/`.
- Determinism comparisons and rollback evidence capture tooling.

5. `Rollback Support Surface`
- Explicit rollback scripts/docs/checklists scoped to WP-20 bridge changes.

## Forbidden Implementation Surface Categories

The following remain forbidden until separately approved:

1. Direct Trio integration behavior.
2. SmartStat engine/apply call activation outside approved bridge contract scope.
3. Viewer implementation behavior (WP-19/UI work) in WP-20 lane.
4. Mutation of captured/validated/projected/adapter artifacts.
5. Unapproved schema/contract changes in WP-17/WP-18/WP-19 surfaces.
6. Unscoped resolver/performance refactors unrelated to bridge scope.
7. Hidden side-effect paths that bypass documented bridge contracts.

## Prototype Branch Merge-Readiness Criteria

Any future WP-20 prototype branch must satisfy all of the following:

1. Charter compliance evidence is complete.
2. Regression/evidence execution plan checkpoints are complete and recorded.
3. Rollback evidence package is complete and reproducible.
4. Protected-surface checks are recorded and reviewed.
5. Determinism and fail-closed evidence are reviewed and accepted.
6. Governance sign-off is recorded before merge.

## Risk Class Statement

WP-20 runtime bridge is a distinct risk-class lane with explicit runtime and
side-effect risk. It requires stricter gating than read-only architecture work.

## Target-02 Outcome

Lane charter is defined.  
WP-20 implementation remains NOT STARTED in this pass.
