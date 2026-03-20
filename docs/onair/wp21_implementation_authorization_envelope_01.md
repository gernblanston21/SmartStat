# WP21 Implementation Authorization Envelope 01

Status date: `2026-03-20`  
Task ID: `WP21_IMPLEMENTATION_AUTHORIZATION_ENVELOPE_01`  
Scope: `Docs-only authorization envelope definition`  
Lane impact: `WP-20 unchanged; WP-21 implementation not started`

## Role Summary

This envelope defines whether one bounded WP-21 implementation step may begin,
without changing runtime behavior, WP-20 payloads, or authorization boundaries.

## WP-21 Implementation Authorization Decision

`AUTHORIZED (CONDITIONALLY)`

## Authorized Objective (ONLY if AUTHORIZED)

Name:
- `WP21_DECISION_MODEL_READONLY_OUTPUT_SURFACE_PASS_01`

Exact purpose:
- Implement one read-only WP-21 decision output surface that consumes existing
  WP-20 preview outputs and emits exactly the WP-21 decision model fields
  already defined in `docs/onair/wp21_decision_layer_definition.md`.

Boundaries:
- one objective only
- advisory-only output
- no runtime execution
- no payload mutation

Implementation host (clarified):
- downstream non-runtime consumer layer only
- outside `SmartStat_v4.1.0.vbs` runtime preview assembly
- output emitted as a separate WP-21 read-only decision surface/artifact
- initial host path reservation:
  `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`
- implementation contract reference:
  `docs/onair/wp21_decision_surface_contract_01.md`

## Non-Overreach Proof

This objective does not duplicate existing WP-20 preview fields:

- `issues_summary` remains detailed evidence (errors/warnings arrays)
- `rule_evaluation_summary_preview` remains detailed rule outcomes and order
- `rule_evaluation_trace_preview` remains traceability/identity evidence
- `ineligible_evidence_preview` remains non-eligibility evidence block

WP-21 introduces only a bounded operator decision summary surface:

- `decision_status`
- `decision_reason`
- `operator_message`
- `recommended_action`

This does not reinterpret runtime semantics:

- mapping rules are fixed and deterministic
- no inferred entities, priorities, or hidden reasoning layers
- no semantic rewriting of WP-20 evidence

## Allowed Surfaces

Only WP-21 decision outputs may be introduced:

- `decision_status`
- `decision_reason`
- `operator_message`
- `recommended_action`

Constraint:
- these are WP-21 outputs only and must not alter WP-20 fields.

## Forbidden Surfaces

The following are forbidden in the authorized objective:

- runtime-adjacent host implementation inside WP-20 preview assembly
- modifications to any WP-20 preview fields, including:
  - `projection_metadata`
  - `issues_summary`
  - `rule_evaluation_summary_preview`
  - `rule_evaluation_trace_preview`
  - `ineligible_evidence_preview`
- any `.vbs` runtime behavior changes
- any mutation/apply/sockets/Trio command paths
- any inference or fallback synthesis layers
- any semantic reinterpretation beyond the fixed WP-21 mapping rules

## Determinism & Failure Guarantees

The authorized objective must preserve:

- deterministic classification output for identical input
- deterministic ordering where output ordering exists
- fail-closed handling for missing/malformed required WP-21 inputs
- no ambiguous output states
- no fallback synthesis for absent evidence

## Implementation Boundary (Future Pass Constraint)

Any future implementation pass under this envelope must:

1. remain read-only and advisory-only
2. consume only existing WP-20 outputs listed in WP-21 direction doc
3. emit only the allowed WP-21 decision output fields
4. avoid new dependencies that alter runtime sequencing or behavior
5. avoid all WP-20 contract/payload modifications
6. conform exactly to `docs/onair/wp21_decision_surface_contract_01.md`

## Risks

- Overreach risk: treating WP-21 summary as execution authority.
- Misclassification risk: implementing mapping outside fixed deterministic rules.
- Operator misuse risk: using advisory output as automatic execution trigger.

## Review

This envelope preserves:

- WP-20 frozen/gated state
- WP-21 non-runtime, non-mutation boundaries
- single-objective discipline
- no scope expansion into implementation beyond one bounded start objective

This envelope does not authorize runtime re-entry or WP-20 modification.

## Real-World Impact

This enables one safe first step: operators get a simple, consistent
"safe/review/blocked" signal from existing preview truth, without changing live
graphics behavior. It improves decision speed while preventing accidental
automation or unsafe execution.

## Non-Authorizing Runtime Boundary

This envelope authorizes only the bounded WP-21 objective above.  
It does not authorize WP-20 runtime re-entry, mutation/apply behavior, or any
runtime `.vbs` code change.

## Authorization Refinement Result (PASS_05)

Decision:
- `IMPLEMENTATION AUTHORIZED`

Bounded implementation scope now allowed:
- exactly one implementation pass for
  `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`
- must conform exactly to
  `docs/onair/wp21_decision_surface_contract_01.md`

Explicitly not authorized in this refinement:
- any `.vbs` runtime edits
- any WP-20 preview payload edits
- any viewer truth-surface expansion
- any additional implementation files unless separately re-authorized
