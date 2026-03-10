# WP-19 Plan Viewer Contract (Read-Only Consumer Layer)

Status: `Target-06 adapter-contract defined` (2026-03-09)  
Scope: docs/tests/tooling only. No viewer UI implementation in this pass.

## Purpose

WP-19 defines a read-only Plan Viewer contract that consumes artifacts produced by:

- WP-17 plan capture contract layer.
- WP-18 validation contract layer.

This contract exists to standardize how a future viewer may inspect plans and validation outcomes without mutating artifacts or invoking runtime behavior.

## Accepted Inputs

WP-19 may consume the following inputs:

1. WP-17 captured-plan artifacts (for example, `contract_version=wp17.plan_capture.v1`).
2. WP-18 validator payload rows containing:
   - `validation_result`
   - `rule_evaluations`
   - refusal diagnostics (`errors`, `warnings`, refusal codes/messages)
   - deterministic identities (`input_identity`, `normalized_plan_hash`, `replay_identity`, `validator_run_identity`)
   - `semantic_interpretation` metadata

## Read-Only Consumption Rules

WP-19 consumption is inspection-only:

1. WP-19 must not mutate captured-plan artifacts.
2. WP-19 must not mutate validation artifacts.
3. WP-19 must preserve ordering already emitted by WP-18 (phase/rule/ordering surfaces are treated as authoritative).
4. WP-19 must not synthesize execution outcomes.
5. WP-19 must not rewrite plan or validation payloads.

## Minimal View-Model Expectations (Contract Scaffold)

WP-19 Target-01 defines minimum read-only view-model surfaces for future implementation:

1. Artifact identity surface:
   - artifact path / fingerprint identity from input artifacts.
2. Validation status surface:
   - PASS/REFUSE summary plus deterministic error/warning lists.
3. Rule evaluation surface:
   - ordered display of rule evaluations by category and declared order.
4. Deterministic identity surface:
   - normalized/replay/validator-run identities displayed as read-only trace fields.
5. Semantic interpretation surface:
   - display of `semantic_interpretation` exactly as validation metadata.

These are contract expectations only; no UI behavior is implemented in this target.

## Target-03 Projection Contract Shape (Read-Only)

WP-19 Target-03 defines deterministic projection fixtures with a stable contract surface:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `status_summary`
6. `issues_summary`
7. `rule_evaluation_summary`
8. `deterministic_identity_summary`
9. `semantic_interpretation_summary`

Projection rules:

1. Projection is derived from WP-17/WP-18 artifacts only.
2. Projection preserves WP-18 rule evaluation ordering.
3. Projection remains read-only and non-mutating.
4. Projection must not introduce runtime/apply/bridge fields.

## Target-04 Projection Summary Contract Hardening (Read-Only)

WP-19 Target-04 hardens the Target-03 projection summary contract by making
summary key shapes and ordering explicit and deterministic.

Top-level projection key order:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `status_summary`
6. `issues_summary`
7. `rule_evaluation_summary`
8. `deterministic_identity_summary`
9. `semantic_interpretation_summary`

Nested summary key order:

1. `input_identity`:
   - `artifact_path`
   - `input_fingerprint_sha256`
2. `status_summary`:
   - `status`
   - `error_count`
   - `warning_count`
3. `issues_summary`:
   - `errors`
   - `warnings`
4. `rule_evaluation_summary`:
   - `phase_order`
   - `ordered_rules`
5. `deterministic_identity_summary`:
   - `normalized_plan_hash`
   - `replay_identity`
   - `validator_run_identity`
6. `semantic_interpretation_summary`:
   - `scope_resolution`
   - `effective_scope`
   - `evidence_source`

Hardening guarantees:

1. Projection summary keys are exact and stable.
2. Projection ordering is deterministic for identical input.
3. Projection preserves WP-18 ordered `rule_evaluations` without re-ordering.
4. Projection remains read-only and non-mutating for source artifacts.
5. Projection remains free of runtime/apply/bridge fields.

## Target-05 Projection Consumption Contract Consolidation (Read-Only)

WP-19 Target-05 consolidates the projection contract into a stable
implementation handoff surface for future viewer work.

### Authoritative Projection Fields (Must Remain Stable)

The following top-level projection fields are authoritative and are consumed as
contract fields in this exact order:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `status_summary`
6. `issues_summary`
7. `rule_evaluation_summary`
8. `deterministic_identity_summary`
9. `semantic_interpretation_summary`

### Display/Summary Surfaces

These fields are viewer display/summary surfaces:

1. `status_summary`
2. `issues_summary`
3. `rule_evaluation_summary`
4. `semantic_interpretation_summary`

Display/summary rules:

1. These fields are derived from WP-18 validation artifacts only.
2. Ordering within these fields must remain deterministic.
3. `rule_evaluation_summary.ordered_rules` must preserve upstream WP-18 order.

### Traceability Surfaces

These fields are traceability and lineage surfaces:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `deterministic_identity_summary`

Traceability rules:

1. Traceability fields are not execution signals.
2. Traceability fields must remain read-only metadata.
3. Traceability fields must support reproducible replay and audit paths.

### Semantic Interpretation Boundaries

`semantic_interpretation_summary` is validation interpretation metadata only:

1. `scope_resolution`
2. `effective_scope`
3. `evidence_source`

Interpretation rules:

1. These values represent validator interpretation, not runtime execution.
2. Omitted stats scope default behavior is represented metadata-only.
3. No projection field may imply apply or bridge behavior.

### Deterministic Ordering Guarantees

1. Top-level field order is stable and contract-bound.
2. Nested summary key order is stable and contract-bound.
3. Rule-evaluation phase order remains `STRUCTURAL`, `SEMANTIC`,
   `DETERMINISM`, `BOUNDARY`.
4. Identical input artifacts must yield identical projection JSON content.

### Explicit Non-Goals

WP-19 projection consumption does not perform:

1. Runtime execution.
2. Apply behavior.
3. Bridge behavior.
4. Artifact mutation.
5. Runtime inference beyond validation artifacts.

## Target-06 Read-Only Projection-to-View-Model Adapter Contract

WP-19 Target-06 defines a strict read-only adapter contract that maps the
consolidated WP-19 projection surface into a future viewer-facing view-model
surface without implementing viewer UI behavior.

### Contract Layers (Separation of Concerns)

1. Projection contract fields:
   - stable upstream WP-19 projection payload fields.
2. Adapter mapping rules:
   - deterministic field mappings from projection fields into view-model
     sections.
3. Future view-model surfaces:
   - read-only viewer-consumable sections derived by adapter mapping only.

### Adapter Contract Identity

1. Adapter contract id: `wp19.projection_to_view_model_adapter.v1`
2. Source projection contract id: `wp19.viewer_projection.v1`
3. Adapter is read-only and deterministic for identical projection input.

### Adapter Output Shape (Read-Only)

Top-level adapter output key order:

1. `adapter_contract`
2. `source_projection_contract`
3. `source_projection_kind`
4. `view_model`

`view_model` key order:

1. `status_view`
2. `issues_view`
3. `rules_view`
4. `semantic_view`
5. `trace_view`

### Strict Adapter Mapping Rules

Status mappings:

1. `status_summary.status` -> `view_model.status_view.status`
2. `status_summary.error_count` -> `view_model.status_view.error_count`
3. `status_summary.warning_count` -> `view_model.status_view.warning_count`

Issues mappings:

1. `issues_summary.errors` -> `view_model.issues_view.errors`
2. `issues_summary.warnings` -> `view_model.issues_view.warnings`

Rules mappings:

1. `rule_evaluation_summary.phase_order` -> `view_model.rules_view.phase_order`
2. `rule_evaluation_summary.ordered_rules` -> `view_model.rules_view.ordered_rules`

Semantic mappings:

1. `semantic_interpretation_summary.scope_resolution` ->
   `view_model.semantic_view.scope_resolution`
2. `semantic_interpretation_summary.effective_scope` ->
   `view_model.semantic_view.effective_scope`
3. `semantic_interpretation_summary.evidence_source` ->
   `view_model.semantic_view.evidence_source`

Trace mappings:

1. `projection_contract` -> `view_model.trace_view.projection_contract`
2. `projection_kind` -> `view_model.trace_view.projection_kind`
3. `input_artifact` -> `view_model.trace_view.input_artifact`
4. `input_identity` -> `view_model.trace_view.input_identity`
5. `deterministic_identity_summary` ->
   `view_model.trace_view.deterministic_identity_summary`

### Deterministic Guarantees

1. Mapping rules are fixed and explicit.
2. Adapter output key order is stable.
3. Adapter output for identical projection input is byte-stable JSON.
4. Adapter preserves projection ordering surfaces (`phase_order`, `ordered_rules`).

### Adapter Non-Goals

The adapter must not perform:

1. Runtime execution.
2. Apply behavior.
3. Bridge behavior.
4. Artifact mutation.
5. Runtime inference beyond projection/validation artifacts.
6. Viewer UI logic or rendering behavior.

## Forbidden Behaviors

WP-19 must not:

1. Execute runtime behavior.
2. Trigger apply behavior.
3. Introduce bridge behavior.
4. Call SmartStat engine/apply surfaces.
5. Introduce Trio integration.
6. Infer runtime side effects from architecture-only artifacts.

## Boundaries and Non-Goals

1. WP-19 Target-01 does not implement viewer UI code.
2. WP-19 Target-01 does not change WP-17/WP-18 logic.
3. WP-19 Target-01 does not activate WP-20.

## Traceability

- Upstream capture contract: `docs/onair/plan-capture-contract.md`
- Upstream validation contract: `docs/onair/plan-validation-contract.md`
- WP-18 validator result model reference: `tests/wp-18/validator/validation_result_model.md`
