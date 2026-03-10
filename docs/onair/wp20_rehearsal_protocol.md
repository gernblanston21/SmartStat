# WP-20 Pre-Implementation Rehearsal Protocol (Target-03)

Status date: `2026-03-10`  
Scope: `WP-20 Target-03 governance/docs/tests rehearsal gate`  
Implementation state: `NOT STARTED`

## Purpose

Define a dry-run rehearsal protocol for a future WP-20 runtime-bridge prototype
lane without implementing runtime-bridge behavior.

The rehearsal validates process readiness, evidence discipline, and fail-closed
governance flow before any implementation is authorized.

## Non-Goals

1. No runtime-bridge code implementation.
2. No apply behavior.
3. No Trio integration.
4. No SmartStat engine/apply calls.
5. No viewer behavior changes.
6. No mutation of captured/validated/projected/adapter artifacts.

## Dry-Run Checkpoint Sequence

Future rehearsal execution must run checkpoints in this order and emit required
outputs for each checkpoint.

1. `R0 - Rehearsal Intake`
- Required outputs:
  - `checkpoint_r0_intake.md`
  - `checkpoint_r0_scope_assertions.json`
- Confirms approved scope, non-goals, and protected surfaces.

2. `R1 - Artifact Pack Skeleton`
- Required outputs:
  - `checkpoint_r1_pack_layout.md`
  - `checkpoint_r1_pack_tree.txt`
- Confirms artifact pack structure is complete and correctly named.

3. `R2 - Checkpoint Mapping`
- Required outputs:
  - `checkpoint_r2_checkpoint_map.md`
  - `checkpoint_r2_required_outputs.json`
- Maps each planned runtime-bridge checkpoint to required evidence artifacts.

4. `R3 - Evidence Naming and Location Validation`
- Required outputs:
  - `checkpoint_r3_naming_validation.md`
  - `checkpoint_r3_location_validation.md`
- Confirms naming/location rules are deterministic and policy-compliant.

5. `R4 - Fail/Abort Rehearsal`
- Required outputs:
  - `checkpoint_r4_fail_abort_matrix.md`
  - `checkpoint_r4_abort_path.md`
- Exercises fail/abort decision flow as documentation-level dry-run.

6. `R5 - Completion Adjudication`
- Required outputs:
  - `checkpoint_r5_completion_gate.md`
  - `checkpoint_r5_rehearsal_result.json`
- Determines `REHEARSAL_COMPLETE` or `REHEARSAL_FAIL`.

7. `R6 - Governance Sign-Off`
- Required outputs:
  - `checkpoint_r6_signoff.md`
  - `rehearsal_index.md`
- Captures final governance decision and evidence index.

## Required Rehearsal Artifact Pack Structure

Future rehearsal runs must emit artifacts under:

- `tests/wp-20/target-03/artifacts/<rehearsal_label>/`

Required pack layout:

1. `00_metadata/`
2. `01_r0_intake/`
3. `02_r1_pack_layout/`
4. `03_r2_checkpoint_map/`
5. `04_r3_naming_location/`
6. `05_r4_fail_abort/`
7. `06_r5_completion/`
8. `07_r6_signoff/`
9. `index/`

No root-level temporary artifacts are allowed.

## Evidence Naming and Location Rules

1. File naming must be lowercase snake_case with checkpoint prefix (`checkpoint_rN_*`).
2. Each checkpoint output must be emitted in its checkpoint directory.
3. A top-level `rehearsal_manifest.json` must list every expected and actual output.
4. Missing required outputs are fail-closed rehearsal failures.
5. Rehearsal labels must be deterministic and sortable (example:
   `wp20_rehearsal_YYYYMMDD_hhmmss`).

## Fail / Abort Handling

Rehearsal must enter `ABORT` if any condition occurs:

1. Scope drifts beyond approved governance documents.
2. Protected surface rules are violated.
3. Required checkpoint outputs are missing or malformed.
4. Naming/location rules are violated.
5. Fail/abort matrix does not define an explicit decision path.

Abort handling requirements:

1. Emit `abort_report.md` with trigger and checkpoint context.
2. Emit `abort_decision.json` with owner and timestamp.
3. Stop rehearsal and mark status `REHEARSAL_FAIL`.

## Completion Criteria

Rehearsal completion requires all of the following:

1. All checkpoints `R0` through `R6` completed.
2. All required outputs present and valid.
3. Fail/abort matrix fully defined and reviewed.
4. Rehearsal result marked `REHEARSAL_COMPLETE`.
5. Governance sign-off recorded.

Rehearsal failure is declared if any completion criterion is unmet.

## Target-03 Outcome

Rehearsal protocol is defined for future dry-runs.  
WP-20 implementation remains NOT STARTED in this pass.
