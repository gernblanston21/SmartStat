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

## Intentionally Missing In Phase 3

Not yet implemented:
- semantic query-path scoring/ranking
- full plan-step trace graph rendering
- deep UI-oriented hierarchy optimization

Those are intentionally deferred to preserve deterministic and evidence-first behavior.
