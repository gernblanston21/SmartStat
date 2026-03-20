# Runtime Slice-02E Resolution Preview Artifact Index

| File Path | Artifact Purpose | Relevance to Closeout |
|---|---|---|
| `docs/onair/wp20-runtime-slice-02/archive/validation_matrix.md` | Consolidated validation outcomes and hashes | Primary status reference for closeout |
| `tools/onair/wp20-runtime-slice-02/run_slice2_fixture.ps1` | Existing fixture execution helper | Reproducible validation invocation surface |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/plan_bridge_case01.fixture` | Positive carry-forward fixture | Slice-02A positive carry-forward verification input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_missing_tabfield_list.fixture` | Existing negative fixture | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_duplicate_tabfields.fixture` | Existing negative fixture | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_unsupported_surface.fixture` | Existing negative fixture | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_contract_noncanonical_pagename.fixture` | Existing contract-hardening negative fixture | Slice-02A prior fail-closed carry-forward evidence |
| `tests/wp-19/target-03/fixtures/projection_pass_case.json` | Projection-intake positive fixture | Source of already-authorized metadata reused by slice-02E |
| `tests/wp-19/target-03/fixtures/projection_refuse_case.json` | Projection-intake negative fixture | Slice-02B runtime-eligibility fail-closed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_unsupported_contract.json` | Projection-intake negative fixture | Slice-02B unsupported-contract fail-closed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_status.json` | Projection-intake negative fixture | Resolution-preview required `status_summary.status` fail-closed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_semantic_scope.json` | Semantic-intake negative fixture | Resolution-preview required `scope_resolution` fail-closed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_empty_semantic_evidence_source.json` | Semantic-intake negative fixture | Resolution-preview required `evidence_source` fail-closed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_missing_issues_errors.json` | Issues-intake negative fixture | Slice-02D carry-forward malformed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/projection_malformed_issues_warnings_not_array.json` | Issues-intake negative fixture | Slice-02D carry-forward malformed evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run1.json` | Slice-02A positive run output (run 1) | Carry-forward deterministic source A |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run2.json` | Slice-02A positive run output (run 2) | Carry-forward deterministic source B |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_missing_tabfield_list.json` | Existing negative output | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_duplicate_tabfields.json` | Existing negative output | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_unsupported_surface.json` | Existing negative output | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_contract_noncanonical_pagename.json` | Existing contract-hardening negative output | Slice-02A prior fail-closed carry-forward evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json` | Resolution-preview positive output (run 1) | Deterministic joined-preview source A |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.json` | Resolution-preview positive output (run 2) | Deterministic joined-preview source B |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_missing_artifact.json` | Projection-intake negative output | Slice-02B missing-artifact fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_unsupported_contract.json` | Projection-intake negative output | Slice-02B unsupported-contract fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_status.json` | Projection-intake negative output | Resolution-preview required `status` malformed fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_not_runtime_eligible.json` | Projection-intake negative output | Slice-02B runtime-eligibility fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_semantic_scope.json` | Semantic-intake negative output | Resolution-preview required `scope_resolution` malformed fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_empty_semantic_evidence_source.json` | Semantic-intake negative output | Resolution-preview required `evidence_source` malformed fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_missing_issues_errors.json` | Issues-intake negative output | Slice-02D carry-forward malformed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_projection_malformed_issues_warnings_not_array.json` | Issues-intake negative output | Slice-02D carry-forward malformed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/mutation_boundary_report.txt` | Mutation boundary report | Confirms `TOTAL_MUTATION_CALLS=0` |
| `SmartStat_v4.1.0.vbs` | Validated runtime surface containing slice-02 gated region | Confirms bounded read-only resolution-preview runtime scope |

