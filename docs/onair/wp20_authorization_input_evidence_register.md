# WP-20 Authorization-Input Evidence Register (Target-09)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only evidence register`  
Implementation state: `NOT STARTED`

## Purpose

Provide a formal register for real authorization-input evidence items that must
be collected and verified before implementation-authorization review.

## Register Fields

Each evidence item must include:

1. `evidence_item_id`
2. `evidence_category`
3. `source_path`
4. `source_owner`
5. `collection_status`
6. `verification_status`
7. `blocker_status`
8. `notes`

## Register Template

| evidence_item_id | evidence_category | source_path | source_owner | collection_status | verification_status | blocker_status | notes |
|---|---|---|---|---|---|---|---|
| T09-EVID-001 | governance_input | `fill_required` | `fill_required` | `not_started` | `not_started` | `none` | `fill_required` |
| T09-EVID-002 | branch_lane | `fill_required` | `fill_required` | `not_started` | `not_started` | `none` | `fill_required` |
| T09-EVID-003 | boundary_assurance | `fill_required` | `fill_required` | `not_started` | `not_started` | `none` | `fill_required` |

## Status Value Rules

- `collection_status`: `not_started | in_progress | complete`
- `verification_status`: `not_started | in_progress | verified | failed`
- `blocker_status`: `none | present | resolved`

## Fail-Closed Rule

If any required evidence item is missing, uncollected, unverified, failed, or
blocked, authorization-input readiness is `NOT READY` and outcome must remain
`hold`.

No missing or unverified evidence may be treated as implicitly acceptable.

## Explicit Non-Authorizing Rule

This register is an input-collection instrument only.  
It does not authorize implementation and does not start WP-20.

