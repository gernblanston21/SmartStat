# Traceability Guide

This guide explains how to trace semantic index records and relationships back to raw OnAir dump evidence.

## Reading Relationship Objects

Each relationship in `relationships[]` provides:
- `id`: deterministic relationship identifier
- `type`: relationship category (for example `profile_to_league`)
- `from_id` / `to_id`: stable semantic record references (or stable synthetic nodes like `league:<code>`)
- `league`: league scope where known
- `source_type`, `source_path`, `source_ref`: provenance anchor for why the relationship exists
- `confidence`: deterministic confidence tier
- `notes`: optional conservative inference notes

## Following Lineage From A Semantic Record

Each semantic record includes:
- `lineage[]`: source lineage entries with `source_type`, `source_path`, `source_ref`, and optional `role`
- `evidence[]`: compact provenance strings derived from lineage
- `relationship_refs[]`: relationship IDs touching this record
- `related_ids[]`: deterministic related record/synthetic IDs

To trace a record:
1. Start at record `id`.
2. Inspect `lineage[]` and `evidence[]` for source provenance.
3. Follow `relationship_refs[]` into `relationships[]`.
4. Use `trace_index.by_source_path` to find all records tied to the same raw source file.

## Plan Engine And React Source View Support

This structure is intended to support:
- Plan Engine explanation/debugging: show why a record exists and which evidence supports it
- Future React SPA Source View: navigate from semantic nodes to source lineage and relationship context

## Query Paths And Lineage

Phase 4 adds `query_paths[]` to provide deterministic entry flows:
- entry record -> league step -> source/lineage step -> terminal references

`query_paths` are built from existing normalized records, `relationships`, and `trace_index` evidence. They do not replace lineage; they pre-compose high-confidence traversal paths for explainable navigation.

## Navigating By League, Record, Source, And Path

Future UI/plan tooling can navigate deterministically from:
1. `league`:
   - use `ui_views.by_league` for record IDs in that league
2. `record`:
   - use `trace_index.by_record_id[record_id]` for lineage/evidence and relationship IDs
3. `source file`:
   - use `trace_index.by_source_path[source_path]` for linked record IDs
4. `query path`:
   - use `query_paths[]` by `entry_record_id`/`path_type` for pre-modeled explanation flows

## How Source Tree And Trace Index Complement Each Other

- `source_tree` is optimized for browse/navigation (source-first hierarchy).
- `trace_index` is optimized for explainability/provenance lookup (record-first and source-first maps).

Together they provide both:
- a UI-friendly source browser model
- a deterministic evidence lookup backbone

## Intentionally Missing In Phase 3

Not yet implemented:
- semantic query-path scoring/ranking
- full plan-step trace graph rendering
- deep UI-oriented hierarchy optimization

Those are intentionally deferred to preserve deterministic and evidence-first behavior.
