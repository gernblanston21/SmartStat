# WP-19 Acceptance (Read-Only Viewer-Contract Package)

Status: CLOSED (accepted 2026-03-09)

## Scope Accepted

WP-19 is accepted as a read-only consumer-contract package over WP-17 and WP-18
artifacts. Acceptance scope is docs/tests/tooling only.

## Final Contract Surfaces

1. Projection contract surface (`wp19.viewer_projection.v1`) with stable top-level and nested ordering.
2. Projection summary hardening for deterministic key shapes/order.
3. Projection consumption contract separating display/summary and traceability surfaces.
4. Projection-to-view-model adapter contract (`wp19.projection_to_view_model_adapter.v1`) with strict deterministic mapping rules.

## Read-Only Boundaries and Non-Goals

WP-19 does not introduce:

1. Viewer UI implementation behavior.
2. Runtime execution behavior.
3. Apply behavior.
4. Bridge behavior.
5. Trio integration.
6. SmartStat engine/apply calls.
7. Artifact mutation.
8. Runtime inference beyond validation/projection artifacts.

## Harness Evidence

1. `tests/wp-19/harness/read_only_intake_contract_test.py`
2. `tests/wp-19/harness/viewer_projection_contract_test.py`
3. `tests/wp-19/harness/projection_consumption_contract_test.py`
4. `tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`

## Target Evidence Paths

1. `tests/wp-19/target-01/`
2. `tests/wp-19/target-02/`
3. `tests/wp-19/target-03/`
4. `tests/wp-19/target-04/`
5. `tests/wp-19/target-05/`
6. `tests/wp-19/target-06/`

## Handoff Boundary to Future Implementation Lane

Future implementation may consume WP-19 projection and adapter contract surfaces
as read-only inputs only. Any runtime/apply/bridge behavior remains out of scope
until WP-20 kickoff and explicit approval.
