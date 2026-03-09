# WP-18 Validator Phase Order

WP-18 validation runner executes deterministic phases in this order:

1. Schema compatibility
2. Structural rules
3. Semantic rules
4. Determinism rules
5. Boundary rules

Phase behavior:

- Schema failure: structural + semantic + determinism + boundary phases are emitted as `WARN` (deterministic skip).
- Structural failure: semantic + determinism + boundary phases are emitted as `WARN` (deterministic skip).
- Semantic failure: determinism + boundary rules still evaluate where meaningful; validation `status=REFUSE`.
- Determinism failure: boundary rules still evaluate where meaningful; prior refusal status remains.
- Boundary checks are represented by both rule evaluations and replay harness assertions (read-only + architecture-only contract checks).
