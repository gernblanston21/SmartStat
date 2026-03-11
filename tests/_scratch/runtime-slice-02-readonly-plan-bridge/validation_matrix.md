# WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE validation matrix

| Fixture | Expected Result | Actual Result | Pass/Fail |
|---|---|---|---|
| plan_bridge_case01.fixture | status=success | status=success; error_code=(empty) | PASS |
| neg_missing_tabfield_list.fixture | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | status=fail_closed; error_code=FIXTURE_COMMAND_MISSING | PASS |
| neg_duplicate_tabfields.fixture | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | status=fail_closed; error_code=AMBIGUOUS_TABFIELD_LIST | PASS |
| neg_unsupported_surface.fixture | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | status=fail_closed; error_code=FIXTURE_LOAD_FAILED | PASS |
| neg_contract_noncanonical_pagename.fixture | status=fail_closed; error_code=SLICE2_CONTRACT_REQUIREMENT_FAILED | status=fail_closed; error_code=SLICE2_CONTRACT_REQUIREMENT_FAILED | PASS |
| plan_bridge_case01.fixture + projection_pass_case.json | status=success | status=success; error_code=(empty) | PASS |
| plan_bridge_case01.fixture + missing projection artifact | status=fail_closed; error_code=SLICE2_PROJECTION_ARTIFACT_MISSING | status=fail_closed; error_code=SLICE2_PROJECTION_ARTIFACT_MISSING | PASS |
| plan_bridge_case01.fixture + projection_unsupported_contract.json | status=fail_closed; error_code=SLICE2_PROJECTION_CONTRACT_UNSUPPORTED | status=fail_closed; error_code=SLICE2_PROJECTION_CONTRACT_UNSUPPORTED | PASS |
| plan_bridge_case01.fixture + projection_malformed_missing_status.json | status=fail_closed; error_code=SLICE2_PROJECTION_ARTIFACT_MALFORMED | status=fail_closed; error_code=SLICE2_PROJECTION_ARTIFACT_MALFORMED | PASS |
| plan_bridge_case01.fixture + projection_refuse_case.json | status=fail_closed; error_code=SLICE2_PROJECTION_NOT_RUNTIME_ELIGIBLE | status=fail_closed; error_code=SLICE2_PROJECTION_NOT_RUNTIME_ELIGIBLE | PASS |
| mutation_boundary_report.txt | TOTAL_MUTATION_CALLS=0 | TOTAL_MUTATION_CALLS=0 | PASS |

Determinism check (positive fixture):
- pos_plan_bridge_run1.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- pos_plan_bridge_run2.json SHA256 = AF124012ED2DA1C62CD5822C9E39FD295FD01BD925B6735FCEA1901AB836C0B7
- hash_match = True

Determinism check (projection-intake joined preview):
- pos_projection_intake_run1.json SHA256 = 96924D6C0964AAD8EC5267DFD838E917A49673DCB4ECAE510EA91BA7EA8ADB1E
- pos_projection_intake_run2.json SHA256 = 96924D6C0964AAD8EC5267DFD838E917A49673DCB4ECAE510EA91BA7EA8ADB1E
- hash_match = True

Gate-OFF parity check (mock-trio normalized):
- method: mock_trio_runner.vbs baseline/successor comparison with identity-only exclusions
- normalized exclusions: TARGET_SCRIPT, SCRIPT_HEADER, SCRIPT_VERSION_CONST, SMARTSTAT_OPERATORDIAG_<volatile_timestamp>
- baseline normalized SHA256 = 80386630636AB0B96C392D5609DE7A5353A532CAF53FA3EEA359FC6D06FE1790
- successor normalized SHA256 = 80386630636AB0B96C392D5609DE7A5353A532CAF53FA3EEA359FC6D06FE1790
- hash_match = True

Boundary check (slice-02 region static scan):
- page:set_property count = 0
- tabfield:set_custom_property count = 0
- sock:send_socket_data count = 0
