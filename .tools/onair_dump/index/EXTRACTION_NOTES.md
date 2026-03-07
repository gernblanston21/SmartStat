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
