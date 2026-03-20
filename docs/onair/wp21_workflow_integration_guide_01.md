# WP21 Workflow Integration Guide 01

Status date: `2026-03-20`  
Scope: `Docs-only workflow integration for WP21 PASS_01`  
Implementation scope: `WP21_DECISION_MODEL_READONLY_OUTPUT_SURFACE_PASS_01`

## Purpose

Define a repeatable operator workflow for using WP21 during live production
without changing runtime behavior, viewer behavior, or execution authority.

## Workflow Placement

WP21 placement in the broadcast flow:

- Build phase: optional preparation only; WP21 is not the primary action point.
- Preview phase: gather/read WP20 preview evidence.
- Pre-take phase: primary WP21 checkpoint (manual advisory decision check).
- Post-take phase: optional retrospective review only; no execution coupling.

Primary placement for WP21:
- pre-take, after preview evidence exists and before operator executes take.

## Operator Workflow Checklist (Primary)

Use this checklist for each page/item where WP21 review is required:

1. Confirm the correct page/item is loaded for preview review.
2. Confirm WP20 preview evidence artifact exists for that page/item.
3. Run `tools/onair/wp21-decision-layer/Build-Wp21DecisionSurface.ps1` with one
   input artifact and one output artifact path.
4. Open the WP21 decision artifact and read:
   - `decision_status`
   - `decision_reason`
   - `operator_message`
   - `recommended_action`
5. Apply the Status -> Action table below exactly.
6. If `REVIEW_REQUIRED` or `BLOCKED`, review source WP20 preview evidence before
   any take decision.
7. Keep final decision operator-owned; WP21 is advisory-only.
8. Continue normal manual Viz Trio workflow after the check is complete.

## Status -> Action Table

| Status | Meaning | Operator Action | Allowed to Proceed |
|---|---|---|---|
| `AUTO_SAFE` | No blockers or review signals detected by bounded WP21 rules. | Continue normal operator flow and maintain standard show-policy checks. | Yes |
| `REVIEW_REQUIRED` | Non-blocking review signals detected (warnings or non-pass outcomes). | Pause, review WP20 preview evidence details, then decide manually. | Conditional |
| `BLOCKED` | Blocking or unsafe condition detected (for example ineligible evidence, errors, malformed/ambiguous required input). | Do not proceed. Escalate through normal editorial/technical path. | No |

## Under Pressure Behavior (Live Show)

When time is tight, prioritize in this order:

1. `decision_status`
2. `recommended_action`
3. `operator_message`
4. `decision_reason` and linked WP20 evidence detail

Safe to shorten:
- for `AUTO_SAFE`, deep-dive evidence review may be shortened when show policy
  allows and no other red flags exist.

Not safe to skip:
- any blocking response (`BLOCKED`)
- required evidence review for `REVIEW_REQUIRED`
- normal show/editorial policy checks

## Operational Boundaries (Critical)

WP21 does NOT:
- trigger takes
- modify tabfields
- execute runtime logic
- override operator decisions
- authorize runtime re-entry
- expand viewer truth surfaces

WP21 remains:
- downstream non-runtime
- advisory-only
- deterministic and fail-closed under the existing contract

## Broadcast Value

Operational impact in live workflow:

- speed: faster pre-take triage with one bounded status signal
- confidence: consistent wording reduces uncertainty under pressure
- reduced hesitation: clear next action for safe/review/block states
- reduced on-air errors: blockers become explicit before manual take decisions

## Authority References

- `docs/onair/wp21_decision_layer_definition.md`
- `docs/onair/wp21_decision_surface_contract_01.md`
- `docs/onair/wp21_decision_surface_adoption_guide_01.md`
- `docs/onair/wp21_operator_value_surface_definition_01.md`
- `docs/onair/wp20_runtime_reentry_authorization_envelope_01.md`
- `docs/onair/wp20_runtime_lane_freeze_summary.md`
- `docs/onair/viewer-readonly-boundary.md`
- `docs/viz-trio/page_editor.md`
- `docs/viz-trio/page_list.md`
- `docs/viz-trio/show_control.md`
- `docs/viz-trio/environment_constraints.md`
