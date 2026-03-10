# WP-20 Authorization-Input Owner Assignment Template (Target-10)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only owner assignment template`  
Implementation state: `NOT STARTED`

## Template Fields

Each row must include:

1. `authorization_input_category`
2. `assigned_owner`
3. `backup_owner`
4. `collection_due_date`
5. `verification_owner`
6. `blocker_escalation_owner`
7. `status`
8. `notes`

## Assignment Template

| authorization_input_category | assigned_owner | backup_owner | collection_due_date | verification_owner | blocker_escalation_owner | status | notes |
|---|---|---|---|---|---|---|---|
| governance_input | `fill_required` | `fill_required` | `YYYY-MM-DD` | `fill_required` | `fill_required` | `not_started` | `fill_required` |
| branch_lane_input | `fill_required` | `fill_required` | `YYYY-MM-DD` | `fill_required` | `fill_required` | `not_started` | `fill_required` |
| boundary_input | `fill_required` | `fill_required` | `YYYY-MM-DD` | `fill_required` | `fill_required` | `not_started` | `fill_required` |

## Status Values

- `not_started`
- `in_progress`
- `ready_for_review`
- `blocked`
- `complete_non_authorizing`

## Fail-Closed Rule

If any required owner assignment field is missing for any required category,
readiness status is `NOT READY` and outcome must remain `hold`.

No missing assignment may be treated as implicitly acceptable.

## Explicit Non-Authorizing Rule

Owner assignment completion does not authorize implementation and does not start
WP-20.

