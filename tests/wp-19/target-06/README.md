# WP-19 Target-06 Read-Only Projection-to-View-Model Adapter Contract

Target-06 scope:

1. Define strict read-only adapter mapping from WP-19 projection contract fields to
   future viewer-facing view-model sections.
2. Keep adapter output deterministic for identical projection input.
3. Keep adapter behavior read-only and non-mutating for source projection artifacts.
4. Reconfirm non-goals: no execution, no apply, no bridge, no runtime inference.

Harness entry point:

- `tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`

Artifacts:

- `tests/wp-19/target-06/artifacts/`
