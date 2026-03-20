# WP21 Operator Value Surface Definition 01

Status date: `2026-03-20`  
Scope: `Docs-only operator-value definition for WP21 PASS_01`  
Implementation scope: `WP21_DECISION_MODEL_READONLY_OUTPUT_SURFACE_PASS_01`  
Operator-value decision: `OPERATOR_VALUE_SURFACE_DEFINED_WITH_LIMITATIONS`

## Purpose

Define the first safe operator-facing value surface for the existing WP21
decision artifact without expanding runtime, viewer, or execution scope.

## Operator Value Surface

Name:
- `Pre-Take Advisory Decision Check (Manual)`

Definition:
- A manual, read-only checkpoint where the operator reviews the WP21 decision
  artifact generated from the current WP20 preview artifact before deciding
  whether to continue normal page flow, pause for review, or escalate.

## Workflow Placement

First safe placement in workflow:
1. generate or load the current WP20 preview artifact
2. run `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1`
3. review the emitted WP21 decision artifact
4. continue manual operator flow in Viz Trio

Placement intent:
- this is a post-build, pre-take advisory checkpoint during review
- it is not a take trigger and not a replacement for preview evidence review

## Operator Action Model

`AUTO_SAFE`:
- meaning: no blockers or review signals were detected by the bounded WP21
  rules
- operator action: continue with normal operator flow while still using existing
  show policy checks

`REVIEW_REQUIRED`:
- meaning: warnings or non-pass outcomes are present
- operator action: pause and inspect WP20 preview evidence details before
  proceeding

`BLOCKED`:
- meaning: ineligible evidence, blocking errors, malformed required input, or
  ambiguous input shape was detected
- operator action: do not proceed; escalate through normal editorial/technical
  path

## Operational Boundary

Operator may rely on:
- deterministic status classification from existing WP20 evidence
- deterministic reason token/message/action mapping

Operator must still verify manually:
- full preview evidence context and broadcast intent
- show-specific timing/editorial constraints
- any page/template readiness details outside WP21 four-field output

WP21 does NOT decide:
- whether to auto-take
- whether to write/update tabfields
- whether runtime/apply behavior should execute
- whether viewer or runtime surfaces should be expanded

## Broadcast Value

Primary value in live workflow:
- faster triage: one bounded status reduces scattered evidence scanning
- confidence support: consistent language for proceed/review/block decisions
- error prevention: clearer blocker visibility before operator action
- hesitation reduction: deterministic advisory guidance before take decisions

## Safest Current Consumption Pattern

Current safest pattern remains docs-guided manual consumption:
- keep WP21 artifact review external to runtime and viewer truth surfaces
- treat WP21 output as advisory summary only
- use WP20 preview evidence as source-of-truth detail

## Non-Authorizing Boundary

This operator-value definition does not authorize:
- WP20 runtime re-entry
- `.vbs` runtime changes
- viewer truth-surface expansion
- Trio/tabfield/apply/socket execution paths
- automatic take authority

Future adoption/integration work, if pursued, still requires a separate bounded
authorization pass.

## Readiness Note

This definition is clear enough for a later bounded adoption/integration
authorization step, but only within current advisory-only and non-runtime
boundaries.

## Authority References

- `docs/onair/wp21_decision_layer_definition.md`
- `docs/onair/wp21_implementation_authorization_envelope_01.md`
- `docs/onair/wp21_decision_surface_contract_01.md`
- `docs/onair/wp21_decision_surface_adoption_guide_01.md`
- `docs/onair/wp20_runtime_reentry_authorization_envelope_01.md`
- `docs/onair/wp20_runtime_lane_freeze_summary.md`
- `docs/onair/viewer-readonly-boundary.md`
- `docs/viz-trio/page_editor.md`
- `docs/viz-trio/page_list.md`
- `docs/viz-trio/show_control.md`
- `docs/viz-trio/environment_constraints.md`
