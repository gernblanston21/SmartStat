# WP-20 Authorization-Input Tracking Ledger (Target-11)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only tracking mechanics`  
Implementation state: `NOT STARTED`

## Purpose

Define a formal ledger structure for tracking future authorization-input
collection progress over time.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine integration work.
3. Does not execute real authorization-input collection in this pass.
4. Does not mutate protected runtime/config/schema/validation surfaces.

## Required Tracking Fields

1. `ledger_item_id`
2. `authorization_input_category`
3. `source_reference`
4. `assigned_owner`
5. `verification_owner`
6. `status`
7. `status_date`
8. `blocker_state`
9. `escalation_owner`
10. `notes`

## Per-Item Status Progression Rules

Allowed progression:

1. `not_started -> in_progress -> pending_verification -> verified`
2. `blocked` may occur from any non-verified state.
3. `verified` is terminal unless a new blocker reopens the item.

Invalid progression must fail closed.

## Blocker / Escalation Tracking Fields

- `blocker_state`: `none | present | escalated | resolved`
- `blocker_id`:
- `blocker_description`:
- `escalation_owner`:
- `escalation_date`:
- `escalation_notes`:

## Fail-Closed Rule (Stale/Missing/Unknown)

If any required field is missing, any status is unknown, or status age exceeds
defined freshness policy without update, ledger outcome remains `hold` and may
not be treated as ready.

No stale/missing/unknown state may be treated as implicitly acceptable.

## Explicit Non-Authorizing Boundary

Tracking ledger completion and tracking updates do not authorize implementation.
Implementation still requires separate explicit recorded authorization decision.

