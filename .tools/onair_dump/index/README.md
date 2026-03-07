# Semantic Index

This folder contains the deterministic semantic index pipeline for `.tools/onair_dump/`.

## Phase Scope

Phase 1 (completed):
- inventory scaffold only
- raw source file discovery and counts
- placeholder semantic arrays

Phase 2 (this pass):
- first normalized extraction pass from runtime, lookup, grammar, and schema dump signals
- semantic arrays expected to be populated:
  - `entities`
  - `measures`
  - `filters`
  - `profiles`
  - `qualifiers` (only if confidently extractable; may remain empty)

## Normalized Meaning In This Pass

In this repo, "normalized" currently means:
- deterministic records and stable IDs
- explicit source provenance (`source_type`, `source_path`, `source_ref`)
- stable merge behavior across duplicate candidates
- no speculative or hidden semantic inference

## Expected Consumers

- semantic-layer tooling
- Plan Engine inspection
- React SPA Source View

## Artifacts

Current output file:
- `semantic_index.json`

Current schema file:
- `semantic_index.schema.json`
