# WP21 Decision Surface Adoption Guide 01

Status date: `2026-03-20`  
Scope: `Docs-only first adoption guidance for WP21 decision surface`  
Implementation scope: `WP21_DECISION_MODEL_READONLY_OUTPUT_SURFACE_PASS_01`

## Purpose

Provide a safe first-use guide for the existing downstream WP21 decision
surface script:

- `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`

This guide is read-only and non-authorizing.

## What The Script Does

`Build-Wp21DecisionSurface.ps1`:

1. reads one serialized WP20 preview artifact JSON file
2. classifies it deterministically using the WP21 contract rules
3. writes one separate WP21 decision artifact JSON file with exactly four fields

It does not modify the input artifact and does not perform runtime execution.

## Input Artifact

The script consumes one input file path (`-InputArtifactPath`) that points to a
single WP20 preview artifact JSON.

Expected input shape is defined in:

- `docs/onair/wp21_decision_surface_contract_01.md`

Branch handling:

- eligible branch evidence under `preview_payload.projection_metadata` and
  `preview_payload.rule_evaluation_summary_preview` /
  `preview_payload.rule_evaluation_trace_preview`
- non-eligible branch evidence under
  `preview_payload.ineligible_evidence_preview`

## Output Artifact

The script emits one output file path (`-OutputArtifactPath`) containing one
JSON object with fixed key order:

1. `decision_status`
2. `decision_reason`
3. `operator_message`
4. `recommended_action`

No additional fields are allowed.

## Field Meanings

`decision_status`:

- `AUTO_SAFE`: no blockers or review signals detected
- `REVIEW_REQUIRED`: non-blocking review signals detected
- `BLOCKED`: do not proceed; blocking condition detected

`decision_reason`:

- deterministic reason token explaining the status

`operator_message`:

- exact human-readable message mapped from `decision_reason`

`recommended_action`:

- bounded advisory action:
  - `PROCEED_WITH_OPERATOR_FLOW`
  - `REVIEW_PREVIEW_EVIDENCE`
  - `DO_NOT_PROCEED_ESCALATE`

## How To Interpret Status

`AUTO_SAFE`:

- preview evidence is clean under the bounded WP21 decision rules
- continue normal operator flow only within existing approved boundaries

`REVIEW_REQUIRED`:

- warnings or non-pass rule outcomes are present
- review source preview evidence before proceeding

`BLOCKED`:

- ineligible evidence, blocking errors, ambiguous shape, or required evidence
  missing/malformed
- do not proceed; escalate per operational process

## Advisory-Only Boundary

This decision artifact is advisory-only:

- no runtime authorization
- no apply authorization
- no Trio/tabfield/socket mutation authority
- no viewer truth-surface expansion authority

## What This Does NOT Authorize

This guide does not authorize:

- changes to `SmartStat_v4.1.0.vbs` or any runtime `.vbs` behavior
- WP20 runtime re-entry (still governed as not authorized)
- viewer integration changes
- execution/mutation workflows

## Safest Current Consumption Pattern

Current safest first adoption is docs-guided manual consumption:

1. generate WP20 preview artifact using the existing approved read-only process
2. run `Build-Wp21DecisionSurface.ps1` with one input artifact and one output
   artifact path
3. review the emitted WP21 decision artifact as an operator/developer advisory
   summary
4. use WP20 preview evidence as source-of-truth detail for review decisions

Do not wire this output into runtime execution or viewer truth surfaces without
separate explicit authorization.

## Authority References

- `docs/onair/wp21_decision_layer_definition.md`
- `docs/onair/wp21_implementation_authorization_envelope_01.md`
- `docs/onair/wp21_decision_surface_contract_01.md`
- `docs/onair/wp20_runtime_reentry_authorization_envelope_01.md`
- `docs/onair/wp20_runtime_lane_freeze_summary.md`
- `docs/onair/viewer-readonly-boundary.md`
