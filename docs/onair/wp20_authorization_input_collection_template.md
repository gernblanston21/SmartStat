# WP-20 Authorization-Input Collection Template (Target-09)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only authorization-input collection`  
Implementation state: `NOT STARTED`

## Purpose

Provide a formal template for collecting real authorization inputs required by
future WP-20 implementation-authorization review.

## Non-Goals

1. Does not authorize implementation.
2. Does not start runtime-bridge code work.
3. Does not approve apply/Trio/engine integration.
4. Does not mutate runtime/config/schema/validation surfaces.

## Required Real Input Categories

1. Governance decision inputs.
2. Branch/lane isolation inputs.
3. Regression/evidence readiness inputs.
4. Rehearsal and sign-off readiness inputs.
5. Boundary/protected-surface assurance inputs.
6. Risk, rollback, and abort readiness inputs.

## Source-of-Truth References (Fill Required)

- `approval_requirements_ref`:
- `lane_charter_ref`:
- `regression_evidence_plan_ref`:
- `rehearsal_protocol_ref`:
- `rehearsal_manifest_template_ref`:
- `gate_review_checklist_ref`:
- `implementation_authorization_record_ref`:
- `branch_approval_record_ref`:
- `authorization_packet_index_ref`:
- `packet_completeness_checklist_ref`:
- `authorization_input_evidence_register_ref`:

## Ownership Fields (Fill Required)

- `collection_owner`:
- `review_owner`:
- `authorization_owner`:
- `revocation_owner`:
- `branch_owner`:

## Status Fields (Fill Required)

- `collection_status`: `not_started | in_progress | complete`
- `verification_status`: `not_started | in_progress | complete`
- `decision_readiness_status`: `not_ready | partial | ready_for_review`
- `overall_status`: `hold | review_pending | complete_non_authorizing`

## Blocking-Item Fields (Fill Required)

- `blocker_present`: `true | false`
- `blocker_id`:
- `blocker_description`:
- `blocker_owner`:
- `blocker_resolution_target_date`:
- `blocker_resolution_status`:

## Explicit Non-Authorizing Rule

Collected authorization inputs are required preparation only.  
Collected inputs do not authorize implementation and do not start WP-20.

