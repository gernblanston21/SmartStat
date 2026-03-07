# Semantic Index

This folder contains the deterministic semantic index pipeline for `.tools/onair_dump/`.

## Phase Scope

Phase 1 (completed):
- inventory scaffold only
- raw source file discovery and counts
- placeholder semantic arrays

Phase 2 (completed):
- first normalized extraction pass from runtime, lookup, grammar, and schema dump signals
- semantic arrays expected to be populated:
  - `entities`
  - `measures`
  - `filters`
  - `profiles`
  - `qualifiers` (only if confidently extractable; may remain empty)

Phase 3 (completed):
- semantic relationship enrichment and source-to-plan traceability
- deterministic relationship generation from high-confidence evidence only
- deterministic lineage and evidence indexing for explainability workflows

Phase 4 (this pass):
- deterministic query-path modeling for semantic record entry flows
- UI-facing source tree preparation from normalized + trace data
- view-model prep for future React SPA Source View, without adding UI code

## Normalized Meaning In This Pass

In this repo, "normalized" currently means:
- deterministic records and stable IDs
- explicit source provenance (`source_type`, `source_path`, `source_ref`)
- stable merge behavior across duplicate candidates
- no speculative or hidden semantic inference

## Relationship Enrichment Meaning

Relationship enrichment means adding explicit, deterministic graph links between already normalized records when a source-backed connection is strong enough to avoid guesswork.

## Traceability Meaning

Traceability means every normalized record and emitted relationship can be followed back to concrete dump evidence (`source_type`, `source_path`, `source_ref`) and lineage metadata suitable for plan/debug explanation.

## Query Path Meaning

A query path is a deterministic, explainable navigation path from a semantic entry record (for example a measure/filter/profile/entity) through league context and source evidence steps.

## UI-Facing Source Tree Meaning

A UI-facing source tree is a deterministic browse structure organized by source type and logical buckets (league/path) with stable `record_ids` so a future React SPA can render semantic/source navigation directly.

## Expected Consumers

- semantic-layer tooling
- Plan Engine explanation/debugging
- React SPA Source View

## Phase 3 Outputs

Relationship categories in this phase include:
- `measure -> filter`
- `profile -> league`
- `entity -> league`
- `record -> source lineage`

## Phase 4 Outputs

Phase 4 introduces deterministic query-path and source-tree structures, including:
- `league -> measure`
- `league -> filter`
- `profile -> measure/filter context`
- `record -> source lineage -> raw file`
- source tree nodes for `grammar` / `grammar_snapshot` / `runtime` / `lookup_index` / `schema` browsing

## Artifacts

Current output file:
- `semantic_index.json`

Current schema file:
- `semantic_index.schema.json`
