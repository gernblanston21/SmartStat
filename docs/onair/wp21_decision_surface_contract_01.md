# WP21 Decision Surface Contract 01

Status date: `2026-03-20`  
Task ID: `WP21_DECISION_SURFACE_CONTRACT_DEFINITION_PASS_04`  
Scope: `Docs-only contract definition for first bounded WP-21 objective`  
Implementation state: `NOT STARTED`

## 1) Role Summary

Define the exact contract for the first WP-21 downstream decision artifact so
future implementation remains read-only, deterministic, advisory-only, and
strictly outside WP-20 runtime assembly.

## 2) Artifact Role

Reserved implementation host:
- `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`

Allowed role:
- read one serialized WP-20 preview artifact
- evaluate deterministic decision rules defined below
- emit one separate serialized WP-21 decision artifact with exactly four fields

Not allowed:
- any `.vbs` runtime modification
- any mutation/apply/Trio/socket/tabfield behavior
- any write-back to the input preview artifact
- any WP-20 payload field modification
- any additional output fields beyond the four authorized WP-21 fields

## 3) Input Contract

Input transport:
- serialized JSON file input only (single artifact per invocation)
- no stdin/pipeline contract in this first objective

Expected source type:
- WP-20 read-only preview artifact emitted by existing slice chain outputs

Input branch shapes (deterministic):
1. Eligible branch shape
- `preview_payload.projection_metadata`
- `preview_payload.rule_evaluation_summary_preview`
- `preview_payload.rule_evaluation_trace_preview`

2. Non-eligible branch shape
- `preview_payload.ineligible_evidence_preview`

Consumed surfaces (existing WP-20 only):
- `issues_summary`
- `rule_evaluation_summary_preview`
- `rule_evaluation_trace_preview`
- `ineligible_evidence_preview`
- `projection_metadata`

Resolution rule for `issues_summary` source:
- if non-eligible branch shape is present:
  - consume `preview_payload.ineligible_evidence_preview.issues_summary`
- else:
  - consume `preview_payload.projection_metadata.issues_summary`

Ambiguity rule:
- if both branch shapes are present in one input artifact, this is ambiguous and
  must fail closed (see section 6).

## 4) Output Contract

Output transport:
- serialized JSON artifact output (single object)

Output object key order (fixed):
1. `decision_status`
2. `decision_reason`
3. `operator_message`
4. `recommended_action`

Allowed fields and values:
- `decision_status`:
  - `AUTO_SAFE`
  - `REVIEW_REQUIRED`
  - `BLOCKED`
- `decision_reason` (deterministic token set):
  - `INELIGIBLE_EVIDENCE_PRESENT`
  - `ERRORS_PRESENT`
  - `REQUIRED_INPUT_MISSING_OR_MALFORMED`
  - `AMBIGUOUS_INPUT_SHAPE`
  - `WARNINGS_PRESENT`
  - `NON_PASS_RULE_OUTCOME_PRESENT`
  - `NO_BLOCKERS_OR_REVIEW_SIGNALS`
- `operator_message`:
  - deterministic human-readable message mapped 1:1 from `decision_reason`
- `recommended_action`:
  - `PROCEED_WITH_OPERATOR_FLOW`
  - `REVIEW_PREVIEW_EVIDENCE`
  - `DO_NOT_PROCEED_ESCALATE`

Required status-action mapping:
- `BLOCKED` -> `DO_NOT_PROCEED_ESCALATE`
- `REVIEW_REQUIRED` -> `REVIEW_PREVIEW_EVIDENCE`
- `AUTO_SAFE` -> `PROCEED_WITH_OPERATOR_FLOW`

Required reason-message mapping (exact strings):
- `INELIGIBLE_EVIDENCE_PRESENT` -> `Preview is not runtime-eligible. Do not proceed.`
- `ERRORS_PRESENT` -> `Blocking errors are present in preview evidence.`
- `REQUIRED_INPUT_MISSING_OR_MALFORMED` -> `Required preview evidence is missing or malformed.`
- `AMBIGUOUS_INPUT_SHAPE` -> `Preview input shape is ambiguous and cannot be classified safely.`
- `WARNINGS_PRESENT` -> `Warnings are present. Review preview evidence before proceeding.`
- `NON_PASS_RULE_OUTCOME_PRESENT` -> `One or more rule outcomes are not PASS. Review evidence.`
- `NO_BLOCKERS_OR_REVIEW_SIGNALS` -> `No blockers or review signals detected in preview evidence.`

## 4A) Invocation Contract

Reserved script:
- `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`

Required invocation inputs:
- one input JSON artifact path
- one output JSON artifact path

Invocation constraints:
- exactly one input artifact per invocation
- output must be written as one JSON object matching section 4
- no side outputs are required by this contract

## 5) Deterministic Classification Order

Evaluation precedence is strict and first-match-wins:

1. `BLOCKED`
- condition 1A: both branch shapes present (ambiguous input)
  - reason: `AMBIGUOUS_INPUT_SHAPE`
- condition 1B: non-eligible branch shape exists
  - reason: `INELIGIBLE_EVIDENCE_PRESENT`
- condition 1C: `issues_summary.errors` count > 0
  - reason: `ERRORS_PRESENT`
- condition 1D: required inputs for current branch missing/malformed
  - reason: `REQUIRED_INPUT_MISSING_OR_MALFORMED`

2. `REVIEW_REQUIRED` (only if no `BLOCKED` condition matched)
- condition 2A: `issues_summary.warnings` count > 0
  - reason: `WARNINGS_PRESENT`
- condition 2B: any
  `rule_evaluation_summary_preview.ordered_rules[*].outcome <> "PASS"`
  - reason: `NON_PASS_RULE_OUTCOME_PRESENT`

3. `AUTO_SAFE` (only if no prior condition matched)
- reason: `NO_BLOCKERS_OR_REVIEW_SIGNALS`

Tie behavior:
- no ties are emitted
- first matched condition in the order above is the deterministic winner

## 6) Fail-Closed Behavior

Transport/read failure (input file missing/unreadable/not parseable JSON):
- fail process explicitly
- emit no partial decision artifact

Parseable but unsupported/malformed/ambiguous shape:
- emit `BLOCKED`
- emit deterministic reason from this set only:
  - `REQUIRED_INPUT_MISSING_OR_MALFORMED`
  - `AMBIGUOUS_INPUT_SHAPE`
- emit `recommended_action=DO_NOT_PROCEED_ESCALATE`

Missing required fields:
- treated as `REQUIRED_INPUT_MISSING_OR_MALFORMED`

Unsupported input shape:
- treated as `REQUIRED_INPUT_MISSING_OR_MALFORMED`

## 7) Host Boundary

This artifact must NOT:
- modify WP-20 runtime outputs
- modify WP-20 preview payload semantics
- expand viewer truth surfaces
- call Trio/runtime/apply/mutation paths
- create hidden inference/scoring/ranking layers
- produce additional decision categories or fields

## 8) Non-Authorizing Boundary

This contract does not authorize runtime re-entry and does not authorize any
`.vbs` implementation work by itself.

WP-20 remains frozen/gated.
WP-21 remains advisory-only and bounded to one objective.

## 9) Implementation Readiness Constraint

Contract readiness result:
- `CONTRACT_DEFINED_READY_FOR_IMPLEMENTATION_AUTHORIZATION`

Interpretation:
- contract precision is sufficient for a future single bounded implementation
  authorization refinement pass
- this document itself is not implementation authorization
