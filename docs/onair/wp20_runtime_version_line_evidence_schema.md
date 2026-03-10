# WP-20 Runtime Version-Line Evidence Schema (Target-15)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only version-line evidence schema`  
Implementation state: `NOT STARTED`

## Purpose

Define the required artifact shape and metadata fields for storing version-line
decision evidence before any future authorized WP-20 runtime lane entry.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not replace Target-05 authorization controls.
5. Does not replace Target-13 lane-entry controls.
6. Does not replace Target-14 decision-record controls.

## Required Artifact Shape

Evidence artifact must include these top-level sections:

1. `identity`
2. `selection`
3. `linkages`
4. `approval_verification`
5. `baseline_confirmation`
6. `completeness`

## Required Identity Fields

1. `identity.evidence_record_id`
2. `identity.wp_target` (must be `WP-20 Target-15`)
3. `identity.record_date`
4. `identity.record_owner`
5. `identity.source_bundle_ref`

## Required Linkage Fields

1. `linkages.target_05.implementation_authorization_record_ref`
2. `linkages.target_05.branch_approval_record_ref`
3. `linkages.target_05.authorization_outcome_ref`
4. `linkages.target_13.runtime_lane_entry_checklist_ref`
5. `linkages.target_13.version_line_decision_checkpoint_ref`
6. `linkages.target_13.protected_file_checkpoint_ref`
7. `linkages.target_14.version_line_decision_record_ref`

## Required Approval/Verification Fields

1. `approval_verification.prepared_by`
2. `approval_verification.reviewed_by`
3. `approval_verification.verified_by`
4. `approval_verification.verification_date`
5. `approval_verification.verification_outcome`
6. `approval_verification.verification_notes`

## Required Selected-Version Field Constraints

1. `selection.selected_version_line` must exist.
2. Allowed values only:
   - `SmartStat_v4.1.0.vbs`
   - `SmartStat_v4.2.0.vbs`
3. `selection.selection_count` must equal `1`.
4. `selection.selection_consistency` must be `true`.
5. Blank, multiple, conflicting, or ambiguous values are invalid.

## Required Frozen-Baseline Confirmation Field

1. `baseline_confirmation.frozen_baseline_file` must equal
   `SmartStat_v4.0.0_beta.vbs`
2. `baseline_confirmation.no_runtime_edit_confirmed` must be `true`
3. `baseline_confirmation.comparison_readiness_confirmed` must be `true`

## Evidence Completeness and Status Vocabulary

Required status fields:

1. `completeness.required_items_total`
2. `completeness.required_items_complete`
3. `completeness.missing_items_count`
4. `completeness.overall_status`

Allowed `completeness.overall_status` values:

1. `incomplete`
2. `ready_for_verification`
3. `verified_complete`
4. `invalid`
5. `hold`

## Fail-Closed Rule (Invalid/Missing/Ambiguous Shape)

If artifact shape is invalid, required fields are missing, or any field is
ambiguous/conflicting, outcome is `hold` and runtime implementation remains
blocked.

No invalid/missing/ambiguous evidence shape may be treated as implicitly
acceptable.

## Explicit Non-Authorizing Boundary

This schema is governance-only and non-authorizing. Evidence shape compliance
alone does not authorize implementation and does not start WP-20 runtime work.
