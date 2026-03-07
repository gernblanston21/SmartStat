# Deterministic Search Contract

Phase 7 locks the Phase 6 search behavior as a reusable inspection contract for `tools/semantic-source-view`.

## Normalized Match vs Source-Only Hint

Normalized match:
- match kinds: `exact_id`, `exact_name`, `exact_alias`, `prefix`, `exact_token`, `substring`, `browse`
- appears in `normalizedResults`
- treated as a first-class semantic record hit

Source-only hint:
- match kinds: `source_ref`, `source_path`, `evidence`
- appears in `sourceOnlyResults` only
- used for debugging source provenance when no normalized semantic record directly matches

## Deterministic Ranking and Tie-Breaks

Primary ranking score:
1. `exact_id`
2. `exact_name`
3. `exact_alias`
4. `prefix`
5. `exact_token`
6. `substring`
7. `source_ref`
8. `source_path`
9. `evidence`
10. `browse` (empty search)

Deterministic tie-break order:
1. higher score first
2. `recordType` ascending
3. `league` ascending (`null` treated as `~`)
4. `id` ascending

## Stabilization Note

This ranking contract is deterministic inspection logic for the semantic UI.
It is not runtime planner logic and should not be treated as SmartStat apply/execution behavior.
