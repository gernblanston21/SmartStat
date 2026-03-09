# WP-18 Validator Phase Order

WP-18 validation runner executes deterministic phases in this order:

1. Schema compatibility
2. Structural rules
3. Semantic rules
4. Determinism rules

Phase behavior:

- Schema failure: structural + semantic + determinism phases are emitted as `WARN` (deterministic skip).
- Structural failure: semantic + determinism phases are emitted as `WARN` (deterministic skip).
- Semantic failure: determinism rules still evaluate where meaningful; validation `status=REFUSE`.
- No mutation, no runtime/apply calls, no Trio integration.
