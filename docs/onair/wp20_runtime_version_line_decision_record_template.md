# WP-20 Runtime Version-Line Decision Record Template (Target-14)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only version-line decision record template`  
Implementation state: `NOT STARTED`

## Purpose

Provide a formal governance decision-record template to select exactly one WP-20
runtime implementation version line before any authorized runtime work begins.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not replace Target-05 implementation authorization controls.
5. Does not replace Target-13 runtime lane-entry controls.

## Decision Record Identity Fields (Required)

1. `decision_record_id`
2. `decision_date`
3. `decision_owner`
4. `review_cycle_reference`
5. `related_wp_target`: `WP-20 Target-14`
6. `status`: `draft | review | approved | rejected`

## Allowed Version-Line Choices (Only)

Exactly one selection is permitted:

1. `SmartStat_v4.1.0.vbs`
2. `SmartStat_v4.2.0.vbs`

No other version-line value is permitted.

## Required Rationale Fields

1. `selected_version_line`
2. `selection_rationale_summary`
3. `regression_comparison_impact_statement`
4. `rollback_impact_statement`
5. `frozen_baseline_preservation_statement` (must reference
   `SmartStat_v4.0.0_beta.vbs`)
6. `risk_notes`

## Required Approval/Sign-Off Fields

1. `prepared_by`
2. `reviewed_by`
3. `approved_by`
4. `approval_date`
5. `approval_notes`

## Required Linkage Fields

### Target-05 Authorization Linkage (Required)

1. `implementation_authorization_record_ref`
2. `branch_approval_record_ref`
3. `authorization_outcome_ref`

### Target-13 Lane-Entry Linkage (Required)

1. `runtime_lane_entry_checklist_ref`
2. `version_line_decision_checkpoint_ref`
3. `protected_file_checkpoint_ref`

## Exactly-One-Selection Rule

The decision record must contain exactly one selected version line from the
allowed choices.

## Fail-Closed Rule (Blank/Multiple/Conflicting/Ambiguous)

If the selection is blank, includes multiple selections, conflicts across
fields, or is ambiguous in any way, the decision record is invalid and runtime
work remains blocked.

No invalid/ambiguous selection state may be treated as implicitly acceptable.

## Explicit Non-Authorizing Boundary

This template is governance-only and non-authorizing. Decision-record completion
alone does not authorize implementation and does not start WP-20 runtime work.
