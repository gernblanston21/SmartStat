# Runtime Slice-02 Validation Artifact Index

| File Path | Artifact Purpose | Relevance to Closeout |
|---|---|---|
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/validation_matrix.md` | Consolidated validation outcomes and hashes | Primary closeout status reference |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/tools/run_slice2_fixture.ps1` | Fixture execution helper for slice-02 gate | Reproducible validation command path |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/plan_bridge_case01.fixture` | Positive deterministic preview input | Gate-ON success/determinism evidence input |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_missing_tabfield_list.fixture` | Missing required command negative input | Fail-closed coverage (`FIXTURE_COMMAND_MISSING`) |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_duplicate_tabfields.fixture` | Duplicate tabfield ambiguity negative input | Fail-closed coverage (`AMBIGUOUS_TABFIELD_LIST`) |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/fixtures/neg_unsupported_surface.fixture` | Unsupported fixture surface negative input | Fail-closed coverage (`FIXTURE_LOAD_FAILED`) |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run1.json` | Positive run output (run 1) | Gate-ON determinism source A |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_plan_bridge_run2.json` | Positive run output (run 2) | Gate-ON determinism source B |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_missing_tabfield_list.json` | Negative output: missing required command | Fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_duplicate_tabfields.json` | Negative output: ambiguous tab list | Fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/neg_unsupported_surface.json` | Negative output: unsupported surface | Fail-closed evidence |
| `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/mutation_boundary_report.txt` | Mutation call scan report | Read-only boundary evidence (`TOTAL_MUTATION_CALLS=0`) |
| `SmartStat_v4.1.0.vbs` | Runtime slice-02 gated scaffold surface under validation | Confirms bounded validated runtime surface |
