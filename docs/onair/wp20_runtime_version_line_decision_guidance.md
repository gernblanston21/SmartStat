# WP-20 Runtime Version-Line Decision Guidance (Target-14)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only version-line decision guidance`  
Implementation state: `NOT STARTED`

## Purpose

Define when and how the WP-20 runtime version-line decision record must be
completed, reviewed, and approved before any authorized runtime lane entry.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not replace Target-05 implementation authorization controls.
5. Does not replace Target-13 runtime lane-entry controls.

## Timing Rule (When Decision Record Must Be Completed)

The decision record must be completed before runtime implementation lane entry
and before any runtime code changes begin.

## Ownership and Review Roles

The following roles should be explicitly assigned:

1. `record_preparer`
2. `technical_reviewer`
3. `governance_reviewer`
4. `authorization_approver`

## Linkage to Target-05 Implementation Authorization

The decision record must be linked to Target-05 artifacts:

1. `docs/onair/wp20_implementation_authorization_record.md`
2. `docs/onair/wp20_branch_approval_record.md`

Decision-record approval must be traceable in the Target-05 authorization
packet context.

## Linkage to Target-13 Runtime Lane-Entry Controls

The decision record must satisfy the version-line decision checkpoint in:

1. `docs/onair/wp20_runtime_implementation_lane_entry_checklist.md`

The chosen line must be unambiguous and consistent with branch/lane separation
and protected-file checkpoints.

## Authorization Boundary Rule

Completion of the decision record alone does not authorize implementation.

Runtime work remains blocked until both:

1. explicit implementation authorization is approved, and
2. runtime lane-entry conditions are satisfied.

## Frozen-Baseline Reminder

`SmartStat_v4.0.0_beta.vbs` remains a frozen protected baseline and must not be
modified by WP-20 runtime implementation targets.

## Explicit Non-Authorizing Boundary

This guidance is governance-only and non-authorizing. WP-20 remains
`NOT STARTED` until explicit authorization and lane-entry requirements are
satisfied.
