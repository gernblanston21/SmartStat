# WP-19 (Plan Viewer) - Contract Scaffolding

WP-19 is a read-only consumer layer over WP-17/WP-18 artifacts.
Status: CLOSED (accepted 2026-03-09; docs/tests/tooling-only package).

Target-01 scope in this directory is scaffolding only:

- No viewer UI implementation.
- No runtime/apply/bridge behavior.
- No SmartStat engine calls.
- No artifact mutation.

## Directory Layout

- `tests/wp-19/artifacts/`  
  Evidence/output placeholder location for future WP-19 harness runs.
- `tests/wp-19/harness/`  
  Read-only intake harness notes and tests.
- `tests/wp-19/target-01/`  
  Target-specific scaffold notes and placeholders for this pass.
- `tests/wp-19/target-02/`  
  Target-specific evidence placeholders for read-only intake contract tests.
- `tests/wp-19/target-03/`  
  Projection contract fixtures and evidence for deterministic read-only view-model projection.
- `tests/wp-19/target-04/`  
  Projection summary contract hardening notes and artifacts for exact key-shape assertions.
- `tests/wp-19/target-05/`  
  Projection consumption contract consolidation notes and handoff artifacts.
- `tests/wp-19/target-06/`  
  Projection-to-view-model adapter contract notes and adapter evidence artifacts.

## Planned Harness Focus (Future Targets)

1. Read-only artifact intake checks (WP-17 + WP-18 inputs).
2. Contract-shape checks for viewer-consumable summary surfaces.
3. Deterministic ordering checks for rule-evaluation presentation.
4. Boundary checks proving no runtime/apply/bridge behavior.
5. Strict projection-to-view-model adapter mapping checks.

## Acceptance Evidence

- Acceptance note: `tests/wp-19/ACCEPTANCE.md`
- Target evidence tree: `tests/wp-19/target-01/` through `tests/wp-19/target-06/`
