# WP-20 Implementation Branch Approval Record (Target-05)

Status date: `2026-03-10`  
Scope: `WP-20 Target-05 implementation-authorization decision gate (docs/tests only)`  
Implementation state: `NOT STARTED`

## Purpose

Define the formal branch approval record format required before any future
WP-20 implementation lane may begin.

This document is governance-only and does not start implementation.

## Non-Goals

1. No runtime-bridge code implementation.
2. No apply behavior implementation.
3. No Trio integration.
4. No SmartStat engine/apply calls.
5. No viewer behavior implementation.
6. No artifact mutation.

## Required Branch Approval Inputs

1. Candidate branch name and base branch.
2. Branch isolation plan and ownership.
3. Allowed implementation surface references (from lane charter).
4. Forbidden surface acknowledgements.
5. Protected-surface integrity expectations.
6. Rollback evidence ownership and escalation route.

## Branch Approval Outcomes

Exactly one outcome is permitted:

1. `authorized_to_start_implementation`
- Branch is approved as the implementation lane, with constraints recorded.

2. `hold`
- Branch plan is incomplete; branch cannot be used until deficiencies are fixed.

3. `denied`
- Branch is rejected for implementation use due to governance/safety failures.

## Ownership and Revocation Expectations

Required owners:

1. `branch_approval_owner`
2. `branch_revocation_owner`
3. `branch_merge_owner`
4. `branch_rollback_owner`

Revocation authority must be explicitly defined before any authorization.

## Branch Approval Record Template

Use section names exactly as listed:

### 1. Record Identity

- `record_version`: `wp20.branch_approval_record.v1`
- `record_label`:
- `record_date_utc`:

### 2. Candidate Branch Definition

- `candidate_branch_name`:
- `base_branch_name`:
- `candidate_branch_head_sha`:
- `branch_purpose`:

### 3. Branch Isolation Assertions

- `separate_lane_asserted`: `true|false`
- `cross_lane_change_policy_ref`:
- `runtime_risk_class_acknowledged`: `true|false`
- `isolation_notes`:

### 4. Scope Constraints

- `allowed_surface_refs`:
- `forbidden_surface_refs`:
- `protected_surface_constraints_ref`:
- `mutation_prohibition_acknowledged`: `true|false`

### 5. Review and Decision Inputs

- `required_input_refs`:
- `input_completeness`: `true|false`
- `input_gaps`:
- `review_notes`:

### 6. Decision

- `decision_outcome`: `authorized_to_start_implementation|hold|denied`
- `decision_rationale`:
- `blocking_conditions`:
- `required_follow_up`:

### 7. Ownership

- `branch_approval_owner`:
- `branch_revocation_owner`:
- `branch_merge_owner`:
- `branch_rollback_owner`:

### 8. Revocation Conditions

- `revocation_triggers`:
- `revocation_path_ref`:
- `revocation_notification_path`:

### 9. Sign-Off

- `final_disposition`: `approved|hold|rejected`
- `signoff_date_utc`:
- `signoff_notes_ref`:

## Relationship to Authorization Record

This branch approval record is a required decision input to:

- `docs/onair/wp20_implementation_authorization_record.md`

Branch approval alone does not authorize implementation start.

## Target-05 Outcome

Implementation branch approval record format is defined.  
WP-20 implementation remains NOT STARTED in this pass.
