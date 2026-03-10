# WP-20 Stop-or-Advance Decision Template (Target-17)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only stop-or-advance decision template`  
Implementation state: `NOT STARTED`

## Purpose

Provide the formal decision template used after governance closeout review to
record whether WP-20 should stop at governance completion or advance to separate
runtime authorization consideration.

## Non-Goals

1. Does not authorize runtime implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not replace Target-05 implementation-authorization controls.
5. Does not replace Target-13 lane-entry controls.

## Required Decision Identity Fields

1. `decision_record_id`
2. `wp_target` (must be `WP-20 Target-17`)
3. `decision_date`
4. `decision_owner`
5. `review_scope_ref`

## Allowed Decision Outcomes (Only)

1. `stop_at_governance_completion`
2. `advance_to_separate_runtime_authorization_consideration`

## Required Rationale Fields

1. `decision_rationale_summary`
2. `criteria_satisfaction_summary`
3. `open_risks_summary`
4. `blocked_conditions_summary`
5. `next_step_recommendation`

## Required Linkage Fields (Governance Evidence)

1. `target_05_authorization_artifacts_ref`
2. `target_13_lane_entry_controls_ref`
3. `target_14_decision_record_ref`
4. `target_15_evidence_checklist_schema_ref`
5. `target_16_review_signoff_ref`
6. `target_17_closeout_criteria_ref`

## Required Frozen-Baseline Confirmation Field

1. `frozen_baseline_file` (must equal `SmartStat_v4.0.0_beta.vbs`)
2. `frozen_baseline_preserved_confirmed` (must be `true`)

## Fail-Closed Rule (Blank/Multiple/Conflicting/Ambiguous Decision)

If decision outcome is blank, multiple, conflicting, or ambiguous, decision
status is `hold` and runtime implementation remains blocked.

No blank/multiple/conflicting/ambiguous decision state may be treated as valid.

## Explicit Rule: Advance Outcome Is Not Implementation Authorization

Choosing `advance_to_separate_runtime_authorization_consideration` does not
authorize runtime implementation. A separate explicit authorization decision is
still required before any runtime code changes begin.

## Explicit Non-Authorizing Boundary

This template is governance-only and non-authorizing. It does not start WP-20
runtime implementation.
