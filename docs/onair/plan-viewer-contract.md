# WP-19 Plan Viewer Contract (Read-Only Consumer Layer)

Status: `Target-01 scaffold` (2026-03-09)  
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
