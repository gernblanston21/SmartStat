# Search Explainability Guide

Phase 6 adds deterministic search ranking and explanation behavior for the Semantic Source View.

## Ranking Rules

Records are scored with explicit deterministic priorities:
1. exact ID match
2. exact name match
3. exact alias match
4. prefix match
5. exact normalized token match
6. substring match in normalized record fields
7. source-only match (`source_ref`, `source_path`, `evidence`)

Stable tie-break order:
1. higher score first
2. record type
3. league
4. record ID

## Match Kinds

Implemented kinds:
- `exact_id`
- `exact_name`
- `exact_alias`
- `prefix`
- `exact_token`
- `substring`
- `source_ref`
- `source_path`
- `evidence`
- `browse` (no search term)

The UI groups kinds into compact labels:
- `exact`
- `prefix`
- `token`
- `substring`
- `source`
- `browse`

## Why Zero-Result Searches Can Still Have Hints

Some terms do not resolve to normalized semantic records directly, but still appear in source-oriented metadata.

When normalized matches are zero, the UI still checks:
- source-only record match channels (`source_ref`, `source_path`, `evidence`)
- `source_tree` labels and node paths
- `trace_index.by_source_path`

This provides deterministic debug hints without inventing new semantics.

## Debugging Value

Search explainability helps developers answer:
- why a record matched
- why another record ranked lower
- why no normalized record matched
- where to inspect next (Source Tree, Traceability, Relationships, Query Paths)

This improves semantic debugging today and supports future Plan Engine explanation UX.

For Phase 8 semantic-resolution scaffolding, see:
- `RESOLUTION_EXPLAINABILITY_GUIDE.md`

## Deferred

- editable search rules
- fuzzy-search libraries
- backend indexing/search services
- runtime SmartStat integration
- plan execution
