# SMARTSTAT_SCRATCH_RETAINED_STATE_MANIFEST_PASS_08

Date: 2026-03-20
Scope: `tests/_scratch/runtime-slice-02-readonly-plan-bridge/`
Mode: Read-only/doc-refresh only
Execution status: Log-only prune is already complete; this manifest records retained state only.

## Current Retained Counts

- Root file count: 113
- `fixtures/` file count: 20
- `runs/` file count: 93
- `runs/*.log` file count: 0

## REFERENCED_KEEP

Retained because they are referenced by tracked files under `docs/onair/`, `docs/ai/`, `tools/onair/`, `SESSION.md`, or `ROADMAP.md`.

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
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/mutation_boundary_report.txt`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_contract_noncanonical_pagename.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_duplicate_tabfields.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_missing_tabfield_list.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_semantic_evidence_source.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_issues_warnings_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_issues_errors.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_semantic_scope.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_status.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_missing_artifact.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_not_runtime_eligible.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_unsupported_contract.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_unsupported_surface.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run1.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run2.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/traceability_subtree_equivalence_report.json`

Basename-only references (kept conservatively): 7
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_empty_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_input_fingerprint_sha256.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_input_identity_artifact_path.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_normalized_plan_hash.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_rule_phase_order.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_validator_run_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_ordered_rules_not_array.json`

## LOCAL_KEEP_FOR_REPLAY

Retained local replay/regression evidence not currently referenced by tracked docs/tools. Intentionally preserved for future revisit/regression needs.

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/pass09_projection_noneligible_missing_semantic_evidence_source.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/pass09_projection_noneligible_structurally_valid.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/.gitkeep`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_fingerprint_sha256.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_input_identity_artifact_path.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_normalized_plan_hash.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_rule_phase_order.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_validator_run_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_ordered_rules_not_array.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_issue_extra_field_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_issue_nonstring_message_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_issue_order_A_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_issue_order_B_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_issue_extra_field.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_issue_nonstring_message.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_issue_order_caseA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_issue_order_caseB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_rule_field_order_caseA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_rule_field_order_caseB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_projection_rule_missing_outcome.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_rule_field_order_A_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_rule_field_order_B_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass03_rule_missing_outcome_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_conflict_input_artifact_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_conflict_replay_identity_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_partial_missing_nested_artifact_path_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_partial_missing_nested_replay_identity_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_projection_conflict_input_artifact.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_projection_conflict_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_projection_partial_missing_nested_artifact_path.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass04_projection_partial_missing_nested_replay_identity.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_conflict_status_vs_rule_outcomes_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_partial_missing_outcome_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_partial_unknown_outcome_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_projection_conflict_status_vs_rule_outcomes.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass05_projection_partial_unknown_outcome.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_conflict_resolution_scope_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_partial_resolution_missing_evidence_source_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_pos_match_resolution_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_projection_conflict_resolution_scope.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_projection_match_resolution_preview.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass06_projection_partial_resolution_missing_evidence_source.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_partial_status_summary_missing_warning_count_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_projection_partial_status_summary_missing_warning_count.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_projection_shadow_status_conflict.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass07_shadow_status_conflict_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_partial_issues_summary_missing_warnings_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_projection_partial_issues_summary_missing_warnings.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_projection_shadow_issues_conflict.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_projection_shadow_rule_summary_conflict.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_shadow_issues_conflict_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass08_shadow_rule_summary_conflict_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_missing_semantic_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_pos_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_pos_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_refuse_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_refuse_runB.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_noneligible_runA.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_pos_regression_run.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_reg_malformed_missing_status.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_reg_missing_projection.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pass09_reg_unsupported_contract.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run_toolcheck.json`

## BLOCKED_PENDING_USER_APPROVAL

Blocked from prune/relocate without explicit user approval.

- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.json`
- `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/slice2_readonly_plan_bridge_outcome.json`

Blocked reason:
- Directly tied to runtime/viewer/test reference surfaces; any move/prune requires explicit approval and coordinated path-update planning.

## Notes

- Log-only prune execution was already completed before this refresh pass.
- `docs/ai/SCRATCH_RUNTIME_SLICE02_RETENTION_POLICY.md` remains valid.
- This manifest does not recommend further deletion.
