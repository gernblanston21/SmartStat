# WP-19 Target-03 Viewer Projection Contract Fixtures

Target-03 scope:

1. Define deterministic read-only projection shapes for WP-19 viewer consumption.
2. Provide fixture coverage for:
   - PASS case
   - REFUSE case
   - stats implicit-default scope case
   - stats explicit scope case
3. Verify projection contract surfaces remain:
   - deterministic
   - read-only
   - non-mutating
   - free of runtime/apply/bridge fields

Harness entry point:

- `tests/wp-19/harness/viewer_projection_contract_test.py`

Fixture location:

- `tests/wp-19/target-03/fixtures/`
