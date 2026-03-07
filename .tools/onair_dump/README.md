# OnAir Dump Dataset

This dataset is a tooling input for semantic-layer development. It captures structured, machine-readable OnAir schema, grammar, lookup, and runtime discovery outputs that can be normalized into a semantic index.

This data is not passive research documentation. It is expected to feed semantic indexing, source inspection tooling, and future React SPA Source View work.

## Inventory

### `_meta/`
This folder appears to contain dataset-level metadata and run descriptors, including endpoint and run index files that describe how the dump was produced.

### `grammar/`
This folder appears to contain league grammar trees and supporting evidence/meta files. It is a primary source for grammar-driven semantic extraction.

### `grammar_snapshot/`
This folder appears to contain point-in-time grammar snapshots by league plus snapshot metadata. It is useful for stable comparisons and change tracking.

### `lookup_index/`
This folder appears to contain lookup index exports by league plus metadata. These files are likely useful for entity/measure lookup normalization.

### `runtime/`
This folder appears to contain captured request/response payloads for runtime endpoints, organized by league. It provides observed runtime surface data for tooling analysis.

### `schema/`
This folder appears to contain schema introspection request/response captures and related schema extracts. It is a foundation for typed semantic modeling.

## Runtime Boundary

This dataset should not be used directly by SmartStat runtime. It is tooling input only and must remain isolated from runtime decision and apply paths.

On feature/semantic-layer, your next commit should probably add:
- .tools/onair_dump/README.md
- a short inventory of:
  - _meta/
  - grammar/
  - grammar_snapshot/
  - lookup_index/
  - runtime/
  - schema/
