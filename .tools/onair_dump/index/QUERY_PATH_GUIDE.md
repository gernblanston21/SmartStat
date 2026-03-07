# Query Path Guide

This guide defines Phase 4 query-path modeling for `.tools/onair_dump/index/semantic_index.json`.

## What Query Paths Are

`query_paths[]` contains deterministic entry flows for semantic records.
Each path starts from one entry record and includes stable steps for league context and source evidence.

Examples:
- `measure_entry`: measure -> league -> source
- `filter_entry`: filter -> league -> source
- `profile_entry`: profile -> league -> source
- `entity_entry`: entity -> league/source context

## Query Paths vs Relationships

- `relationships[]` expresses pairwise graph links (`from_id` -> `to_id`).
- `query_paths[]` expresses a deterministic, ordered traversal sequence for explanation/browsing.

Relationships are graph primitives; query paths are explainable navigation routes composed from those primitives plus lineage evidence.

## Plan Engine Explanation Value

Query paths make it easier to explain:
- why a record appears
- what league context it belongs to
- which raw source evidence supports it

This improves future plan/debug narratives without introducing runtime behavior changes.

## React Source View Preparation Value

Query paths prepare a future React SPA Source View by providing:
- precomputed path IDs
- stable step ordering
- terminal record sets
- direct evidence references

This reduces UI-side inference and keeps view logic deterministic.

## Intentionally Deferred Until UI Work

Deferred in Phase 4:
- interactive query/path graph editing
- user-authored path composition
- UI state models (selection, expansion, focus, virtualized rendering)

These remain deferred until React skeleton work begins.
