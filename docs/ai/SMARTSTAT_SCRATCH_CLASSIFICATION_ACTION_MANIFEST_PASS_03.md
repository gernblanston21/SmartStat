# SMARTSTAT_SCRATCH_CLASSIFICATION_ACTION_MANIFEST_PASS_03

Superseded: This manifest is superseded by `docs/ai/SMARTSTAT_SCRATCH_RETAINED_STATE_MANIFEST_PASS_08.md` (post-log-prune retained-state baseline).
Status: Historical record only. Do not use this file as the current retained-state source of truth.

Date: 2026-03-19
Scope: `tests/_scratch/runtime-slice-02-readonly-plan-bridge/`
Mode: Read-only audit + manifest writing only
Execution status: No deletions, no moves, no staging, no commits

## Why No Execution Is Performed In This Pass

This pass is intentionally documentation-only. Files in this `_scratch` subtree may still be needed for future runtime regressions, re-validation, or audit replay. To preserve safety, this manifest records potential future actions only and applies none of them.

## KEEP_LOCAL_REFERENCE (retained local evidence)

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_contract_noncanonical_pagename.fixture`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_duplicate_tabfields.fixture`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_missing_tabfield_list.fixture`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_unsupported_surface.fixture`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/plan_bridge_case01.fixture`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_empty_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_empty_semantic_evidence_source.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_issues_warnings_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_input_fingerprint_sha256.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_input_identity_artifact_path.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_issues_errors.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_normalized_plan_hash.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_rule_phase_order.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_semantic_scope.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_status.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_validator_run_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_ordered_rules_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_unsupported_contract.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/.gitkeep`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/mutation_boundary_report.txt`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_contract_noncanonical_pagename.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_duplicate_tabfields.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_missing_tabfield_list.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_semantic_evidence_source.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_issues_warnings_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_fingerprint_sha256.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_identity_artifact_path.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_issues_errors.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_normalized_plan_hash.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_rule_phase_order.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_semantic_scope.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_status.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_validator_run_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_ordered_rules_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_missing_artifact.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_not_runtime_eligible.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_unsupported_contract.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_unsupported_surface.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run_toolcheck.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run1.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run2.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/traceability_subtree_equivalence_report.json`

## PRUNE_CANDIDATE (possible future prune - not approved yet)

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_contract_noncanonical_pagename.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_duplicate_tabfields.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_missing_tabfield_list.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_replay_identity.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_semantic_evidence_source.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_issues_warnings_not_array.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_fingerprint_sha256.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_identity_artifact_path.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_issues_errors.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_normalized_plan_hash.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_rule_phase_order.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_semantic_scope.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_status.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_validator_run_identity.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_ordered_rules_not_array.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_missing_artifact.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_not_runtime_eligible.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_unsupported_contract.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_unsupported_surface.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run_toolcheck.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run1.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run2.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.log`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/slice2_readonly_plan_bridge_diag.log`

## RELOCATE_WITHIN_REPO (proposed target paths only)

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/contract_hardening_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/contract_hardening_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/contract_hardening_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/contract_hardening_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/contract_hardening_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/contract_hardening_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/issues_intake_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/issues_intake_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/issues_intake_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/issues_intake_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/issues_intake_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/issues_intake_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/projection_intake_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/projection_intake_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/projection_intake_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/projection_intake_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/projection_intake_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/projection_intake_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/resolution_preview_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/resolution_preview_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/resolution_preview_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/resolution_preview_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/resolution_preview_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/resolution_preview_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/semantic_intake_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/semantic_intake_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/semantic_intake_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/semantic_intake_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/semantic_intake_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/semantic_intake_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/validation_artifact_index.md` -> `docs/onair/wp20-runtime-slice-02/archive/validation_artifact_index.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/validation_closeout.md` -> `docs/onair/wp20-runtime-slice-02/archive/validation_closeout.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/validation_decision.md` -> `docs/onair/wp20-runtime-slice-02/archive/validation_decision.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/validation_matrix.md` -> `docs/onair/wp20-runtime-slice-02/archive/validation_matrix.md`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/tools/.gitkeep` -> `tools/onair/wp20-runtime-slice-02/.gitkeep`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/tools/check_slice2_traceability_drift.ps1` -> `tools/onair/wp20-runtime-slice-02/check_slice2_traceability_drift.ps1`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/tools/run_slice2_fixture.ps1` -> `tools/onair/wp20-runtime-slice-02/run_slice2_fixture.ps1`

## NEEDS_USER_DECISION (blocked)

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/slice2_readonly_plan_bridge_outcome.json`

Blocked reason:
- These are directly referenced by current runtime/viewer/test surfaces; any move/prune requires explicit user authorization and coordinated path-update planning.

## Future Step Order (optional, not executed here)

1. Optional user-decision review for `NEEDS_USER_DECISION` files.
2. Optional prune pass for `PRUNE_CANDIDATE` files only.
3. Optional relocate pass for `RELOCATE_WITHIN_REPO` files using proposed target paths.
4. Optional post-action verification pass (`git status`, reference search, deterministic smoke checks).

## Rollback / Safety Notes

- No files were changed, moved, deleted, staged, or committed in this pass.
- This manifest is advisory only and preserves all current artifacts.
- Any future execution pass should be explicit, bounded, and reversible.
