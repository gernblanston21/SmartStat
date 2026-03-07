# Phase 2 Extraction Notes

## What Was Extracted

Phase 2 upgrades the index from inventory-only to first normalized semantic extraction.

Current populated arrays:
- `profiles`: extracted from `runtime/profiles/profiles.response.json`
- `filters`: extracted from runtime `availableGamePlanFilters` plus lookup split/time maps
- `measures`: extracted from runtime `fetchBaseMeasures` with lookup alias/provenance enrichment
- `entities`: extracted from high-confidence schema type signals

`qualifiers` remains intentionally empty in this pass because no standalone qualifier object shape is strongly represented in current dumps.

## Deferred / Placeholder Work

Deferred in this phase:
- deep qualifier extraction from complex grammar/runtime context
- full relationship modeling between entities, filters, and measures
- broad schema-wide entity synthesis beyond high-confidence allowlisted types

## Known Limitations

- Measure extraction is intentionally conservative in breadth (`MEASURE_LIMIT_PER_LEAGUE`) to keep this first normalized pass deterministic and reviewable.
- Filter extraction is intentionally conservative in breadth (`FILTER_LIMIT_PER_LEAGUE`) for the same determinism-first reason.
- Grammar tree dumps currently expose minimal node information in this dataset, so grammar-driven entity extraction is deferred.
- Duplicate records are merged by stable ID with provenance notes, but semantic conflict resolution is intentionally minimal in Phase 2.

## Likely Next Phase

The next likely step before React SPA work is:
- semantic relationship enrichment
- source-to-plan traceability
- React SPA Source View consuming `semantic_index.json`

---

# Phase 3 Relationship + Traceability Notes

## What Was Added In Phase 3

Phase 3 adds deterministic relationship and traceability structures on top of normalized records.

Added structures:
- top-level `relationships`
- top-level `trace_index`
- per-record lineage/evidence fields for explainability
- per-record relationship references (`related_ids`, `relationship_refs`)

Relationship types emitted in this pass:
- `measure_to_filter` (strict token-match evidence only)
- `profile_to_league`
- `entity_to_league`
- `record_to_source`

## Conservative Inference Policy

This pass remains intentionally conservative:
- relationships are emitted only when source evidence is explicit or string-token linkage is strong
- no broad domain-logic assumptions are introduced
- ambiguous links are omitted rather than guessed

## Deferred In Phase 3

Still deferred:
- deeper semantic dependency inference between measure categories and filter semantics
- cross-record relationship scoring beyond deterministic confidence tiers
- richer qualifier graph extraction from weak/implicit signals

## Next Likely Phase Before React Work

The next likely step is:
- semantic query-path modeling
- plan-oriented trace rendering
- UI-facing source tree preparation
