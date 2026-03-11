# WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE validation matrix

| Fixture | Expected Result | Actual Result | Pass/Fail |
|---|---|---|---|
| plan_bridge_case01.fixture | status=success | status=success; error_code=(empty) | PASS |
| neg_missing_tabfield_list.fixture | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | PASS |
| neg_duplicate_tabfields.fixture | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | PASS |
| neg_unsupported_surface.fixture | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | PASS |
| neg_contract_noncanonical_pagename.fixture | status=fail_closed; error_code=SLICE2_CONTRACT_REQUIREMENT_FAILED | status=fail_closed; error_code=SLICE2_CONTRACT_REQUIREMENT_FAILED | PASS |
| mutation_boundary_report.txt | TOTAL_MUTATION_CALLS=0 | TOTAL_MUTATION_CALLS=0 | PASS |

Determinism check (positive fixture):
- pos_plan_bridge_run1.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- pos_plan_bridge_run2.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- hash_match = True

Gate-OFF parity check (mock-trio normalized):
- method: mock_trio_runner.vbs baseline/successor comparison with identity-only exclusions
- normalized exclusions: TARGET_SCRIPT, SCRIPT_HEADER, SCRIPT_VERSION_CONST, SMARTSTAT_OPERATORDIAG_<volatile_timestamp>
- baseline normalized SHA256 = 80386630636AB0B96C392D5609DE7A5353A532CAF53FA3EEA359FC6D06FE1790
- successor normalized SHA256 = 80386630636AB0B96C392D5609DE7A5353A532CAF53FA3EEA359FC6D06FE1790
- hash_match = True
