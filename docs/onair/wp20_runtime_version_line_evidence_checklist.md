# WP-20 Runtime Version-Line Evidence Checklist (Target-15)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only version-line evidence checklist`  
Implementation state: `NOT STARTED`

## Purpose

Define required evidence checks proving the WP-20 runtime version-line decision
was reviewed correctly, recorded correctly, and linked correctly before any
future authorized runtime lane begins.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not replace Target-05 authorization controls.
5. Does not replace Target-13 lane-entry controls.
6. Does not replace Target-14 decision-record template controls.

## Required Evidence Items - Decision Record Completion

1. Evidence that the Target-14 decision record is present.
2. Evidence that all required identity fields are completed.
3. Evidence that all required rationale fields are completed.
4. Evidence that required approval/sign-off fields are completed.
5. Evidence that the decision record status is final for governance review.

## Required Evidence Items - Exactly-One Allowed Version Selection

1. Evidence that selected version is exactly one value.
2. Evidence that selected value is one of:
   - `SmartStat_v4.1.0.vbs`
   - `SmartStat_v4.2.0.vbs`
3. Evidence that no conflicting version values appear across artifacts.
4. Evidence that no blank/placeholder/ambiguous selection remains.

## Required Evidence Items - Mandatory Linkages

### Target-05 Authorization Linkage

1. Evidence link to implementation-authorization record.
2. Evidence link to branch-approval record.
3. Evidence link to authorization outcome/reference.

### Target-13 Lane-Entry Linkage

1. Evidence link to runtime lane-entry checklist instance.
2. Evidence for version-line decision checkpoint satisfaction.
3. Evidence for protected-file checkpoint satisfaction.

### Target-14 Decision Record Linkage

1. Evidence link to decision-record template instance.
2. Evidence that the instance fields satisfy Target-14 rules.

## Required Evidence Items - Frozen Baseline Preservation

1. Evidence confirming `SmartStat_v4.0.0_beta.vbs` remains frozen/protected.
2. Evidence confirming no WP-20 runtime implementation edits target
   `SmartStat_v4.0.0_beta.vbs`.
3. Evidence confirming baseline remains usable for regression/rollback/governance
   comparison.

## Verification and Owner Fields (Required)

1. `evidence_owner`
2. `verification_owner`
3. `review_owner`
4. `verification_date`
5. `verification_outcome`
6. `verification_notes`

## Fail-Closed Rule (Missing/Incomplete/Conflicting Evidence)

If required evidence is missing, incomplete, conflicting, or ambiguous, evidence
outcome is `hold` and runtime implementation remains blocked.

No missing/incomplete/conflicting evidence state may be treated as implicitly
acceptable.

## Explicit Non-Authorizing Boundary

Checklist completion is governance-only and non-authorizing. Evidence completion
alone does not authorize implementation and does not start WP-20 runtime work.
