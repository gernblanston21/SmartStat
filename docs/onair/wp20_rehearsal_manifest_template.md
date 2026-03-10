# WP-20 Rehearsal Manifest Template (Target-04)

Status date: `2026-03-10`  
Scope: `WP-20 Target-04 pre-implementation sign-off gate (docs/tests only)`  
Implementation state: `NOT STARTED`

## Purpose

Define a formal manifest template for future WP-20 rehearsal runs so evidence
is captured in a deterministic, review-ready format before any implementation
lane can be authorized.

This template does not authorize implementation.

## Non-Goals

1. No runtime-bridge code implementation.
2. No apply behavior.
3. No Trio integration.
4. No SmartStat engine/apply calls.
5. No viewer behavior changes.
6. No artifact mutation.

## Manifest Usage Rule

Future rehearsal runs must emit one manifest file:

- `tests/wp-20/target-03/artifacts/<rehearsal_label>/rehearsal_manifest.md`

The manifest must be complete before gate review begins.

## Manifest Template

Use the following section structure exactly (section names are mandatory).

### 1. Manifest Identity

- `manifest_version`: `wp20.rehearsal_manifest.v1`
- `rehearsal_label`:
- `created_utc`:
- `updated_utc`:
- `owner`:
- `branch`:
- `commit_sha`:

### 2. Scope Assertions

- `approved_scope_refs`:
  - `docs/onair/wp20_approval_requirements.md`
  - `docs/onair/wp20_lane_charter.md`
  - `docs/onair/wp20_regression_evidence_plan.md`
  - `docs/onair/wp20_rehearsal_protocol.md`
- `non_goals_asserted`: `true|false`
- `protected_surfaces_asserted`: `true|false`
- `notes`:

### 3. Checkpoint Results

List checkpoint outcomes in protocol order:

1. `R0` rehearsal intake
2. `R1` artifact pack skeleton
3. `R2` checkpoint mapping
4. `R3` naming/location validation
5. `R4` fail/abort rehearsal
6. `R5` completion adjudication
7. `R6` governance sign-off preparation

For each checkpoint include:

- `status`: `pass|hold|fail`
- `required_outputs_present`: `true|false`
- `artifact_refs`:
- `review_notes`:

### 4. Evidence Pack Inventory

- `artifact_pack_root`:
- `artifact_pack_tree_ref`:
- `required_files_present`: `true|false`
- `missing_files`:
- `naming_policy_compliant`: `true|false`
- `location_policy_compliant`: `true|false`

### 5. Boundary and Protected Surface Verification

- `protected_diff_command`:
- `protected_diff_output_ref`:
- `protected_surfaces_unchanged`: `true|false`
- `runtime_apply_bridge_changes_detected`: `true|false`
- `mutation_detected`: `true|false`

### 6. Failure / Abort Register

- `abort_triggered`: `true|false`
- `abort_reason`:
- `abort_checkpoint`:
- `abort_report_ref`:
- `abort_decision_ref`:

If `abort_triggered=true`, gate outcome must be `fail`.

### 7. Reviewer Sign-Off Inputs

Required role entries:

1. `lane_owner`
2. `governance_reviewer`
3. `determinism_reviewer`
4. `boundary_safety_reviewer`
5. `release_owner`

For each role include:

- `review_status`: `approved|hold|rejected`
- `review_date`:
- `review_comments_ref`:

### 8. Gate Outcome Recommendation

- `recommended_gate_outcome`: `implementation_ready|hold|fail`
- `recommendation_rationale`:
- `blocking_items`:
- `follow_up_actions`:

## Outcome Rule

Allowed values:

1. `implementation_ready`
2. `hold`
3. `fail`

Final gate decision is recorded in
`docs/onair/wp20_gate_review_checklist.md` review output, not in this template
alone.

## Target-04 Outcome

Formal rehearsal manifest template is defined.  
WP-20 implementation remains NOT STARTED in this pass.
