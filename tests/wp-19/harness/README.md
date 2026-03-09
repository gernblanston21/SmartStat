# WP-19 Harness (Read-Only Contract Tests)

Target-02 now adds a read-only intake contract harness:

- `tests/wp-19/harness/read_only_intake_contract_test.py`
- `tests/wp-19/harness/viewer_projection_contract_test.py`

Future WP-19 harnesses should validate:

1. WP-17/WP-18 input-contract compatibility for read-only viewer consumption.
2. Stable and deterministic projection of validation status/errors/warnings/rule order.
3. Read-only boundary enforcement (no input mutation, no runtime/apply/bridge behavior).
4. Exact projection summary key shapes and deterministic key ordering.

Run command:

- `python tests/wp-19/harness/read_only_intake_contract_test.py`
- `python tests/wp-19/harness/viewer_projection_contract_test.py`

No viewer UI harness is introduced in this target.
