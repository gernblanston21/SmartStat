# WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE validation matrix

| Fixture | Expected Result | Actual Result | Pass/Fail |
|---|---|---|---|
| plan_bridge_case01.fixture | status=success | status=success; error_code=(empty) | PASS |
| neg_missing_tabfield_list.fixture | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | PASS |
| neg_duplicate_tabfields.fixture | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | PASS |
| neg_unsupported_surface.fixture | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | PASS |
| mutation_boundary_report.txt | TOTAL_MUTATION_CALLS=0 | TOTAL_MUTATION_CALLS=0 | PASS |

Determinism check (positive fixture):
- pos_plan_bridge_run1.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- pos_plan_bridge_run2.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- hash_match = True
