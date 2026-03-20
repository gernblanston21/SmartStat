# WP-21 Decision Layer Definition

Status date: `2026-03-20`  
Scope: `Docs-only architectural definition`  
Implementation state: `NOT STARTED`  
Authorization state: `IMPLEMENTATION AUTHORIZED FOR ONE BOUNDED OBJECTIVE ONLY (see envelope)`

## 1) Role Summary

WP-21 defines a deterministic, read-only Decision Layer that consumes completed
WP-20 preview outputs and emits bounded operator guidance without adding data,
inference, mutation, or execution behavior.

## 2) WP-21 Core Definition

One-sentence definition:
WP-21 is a deterministic decision-classification surface over existing WP-20
runtime preview truth.

Expanded definition:
WP-21 transforms existing WP-20 preview fields into a small, explicit decision
model for operators. It does not alter WP-20 payloads, does not reinterpret
runtime semantics, and does not trigger any action path.

## 3) Architectural Position

- Upstream: WP-20 read-only runtime preview layer (truth source).
- WP-21: read-only decision classification and operator message formatting.
- Downstream: Viz Trio operator workflow consumption only (human decision).
- Future execution layers: remain separate and unauthorized by WP-21.
- Implementation host for first WP-21 step: downstream non-runtime consumer
  surface (not inside WP-20 runtime preview assembly).
- First objective contract reference:
  `docs/onair/wp21_decision_surface_contract_01.md`

## 4) Inputs (From WP-20)

WP-21 consumes only existing preview surfaces:

- `issues_summary`
- `rule_evaluation_summary_preview`
- `rule_evaluation_trace_preview`
- `ineligible_evidence_preview`
- `projection_metadata`

## 5) Outputs (WP-21 Decision Model)

```json
{
  "decision_status": "AUTO_SAFE|REVIEW_REQUIRED|BLOCKED",
  "decision_reason": "string",
  "operator_message": "string",
  "recommended_action": "PROCEED_WITH_OPERATOR_FLOW|REVIEW_PREVIEW_EVIDENCE|DO_NOT_PROCEED_ESCALATE"
}
```

Required fields:

- `decision_status` enum only:
  - `AUTO_SAFE`
  - `REVIEW_REQUIRED`
  - `BLOCKED`
- `decision_reason`: short deterministic reason token/message.
- `operator_message`: human-readable read-only message.
- `recommended_action` enum only:
  - `PROCEED_WITH_OPERATOR_FLOW`
  - `REVIEW_PREVIEW_EVIDENCE`
  - `DO_NOT_PROCEED_ESCALATE`

## 6) Decision Classification Rules

Rules are evaluated deterministically in this order:

1. `BLOCKED`
- Trigger if `ineligible_evidence_preview` exists.
- Trigger if `issues_summary.errors` count is greater than `0`.
- Trigger if required WP-21 input surfaces are absent or malformed at
  consumption time.
- Emit `recommended_action=DO_NOT_PROCEED_ESCALATE`.

2. `REVIEW_REQUIRED`
- Trigger if `issues_summary.warnings` count is greater than `0` and rule 1 did
  not trigger.
- Trigger if any `rule_evaluation_summary_preview.ordered_rules[*].outcome` is
  not `PASS` and rule 1 did not trigger.
- Emit `recommended_action=REVIEW_PREVIEW_EVIDENCE`.

3. `AUTO_SAFE`
- Trigger only if:
  - rule 1 did not trigger
  - rule 2 did not trigger
  - required inputs are present and parseable
- Emit `recommended_action=PROCEED_WITH_OPERATOR_FLOW`.

No additional rule tiers are allowed in this definition.

## 7) Viz Trio Integration Model

Definition-only (no implementation in this pass):

- WP-21 output may be surfaced in read-only operator UI/script messaging.
- WP-21 output may be logged for operator review.
- WP-21 must not write tabfields, call apply paths, or invoke mutation commands.
- WP-21 must remain advisory and non-authorizing for execution.

## 8) Operator Experience Model

Operator sees one explicit state:

- `AUTO_SAFE`: preview is clean under existing checks.
- `REVIEW_REQUIRED`: warnings/non-pass rule outcomes need manual review.
- `BLOCKED`: do not proceed; resolve blocking conditions first.

This reduces interpretation burden by turning scattered preview evidence into one
deterministic read-only decision summary.

## 9) Boundaries

WP-21 does not:

- create new data
- infer hidden meaning
- reinterpret runtime outputs
- execute runtime/apply behavior
- mutate Viz Trio/tabfields/sockets/engine state
- authorize future runtime implementation by itself

## 10) Risks

- Misclassification risk if input contract is consumed inconsistently.
- Overreach risk if WP-21 is treated as an execution authority.
- Operator confusion risk if messages drift from deterministic rule outputs.

## 11) Real-World Impact

WP-21 gives operators a clear "safe / review / blocked" decision from existing
runtime truth, reducing live-show hesitation and error-prone manual judgment
without changing underlying runtime behavior.

## Non-Authorizing Boundary

This definition is architectural only. It does not modify WP-20, does not start
runtime implementation, and does not authorize mutation/apply behavior.

Governance acceptance of this document (if recorded) is direction-only and does
not grant broad implementation authorization.

Bounded implementation, if later approved, must remain constrained to:
- `WP21_DECISION_MODEL_READONLY_OUTPUT_SURFACE_PASS_01`
- `docs/onair/wp21_implementation_authorization_envelope_01.md`
- `docs/onair/wp21_decision_surface_contract_01.md`
