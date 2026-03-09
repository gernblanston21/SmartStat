# OnAir Semantic Reference Layer

## Purpose / Overview

This directory contains read-only semantic reference artifacts extracted from the OnAir v3 workbook references.

## Source material location

- `docs/onair/source/MLB_OnAir_v3_Stat_Syntax_v3.xlsx`
- `docs/onair/source/NHL_OnAir_v3_Stat_Syntax_v3.xlsx`

## Why this documentation exists

The extracted references provide deterministic source-grounded inputs for:
- semantic dictionaries
- qualifier alias normalization
- candidate-resolution explainability
- future SmartStat plan grammar and Plan Engine preparation

## Included artifacts

- `workbook-inventory.md`
- `measure-dictionary.md`
- `filter-grammar-dictionary.md`
- `alias-map.qualifiers.md`
- `entity-dictionary.md`
- `attribute-dictionary.md`
- `formatter-dictionary.md`
- `query-skeletons.md`
- `onair_semantic_grammar.md`
- `slot-resolution-model.md` (grammar-to-candidate-resolution bridge for future plan capture architecture)
- `plan-capture-shape.md` (conceptual deterministic captured-plan shape after slot resolution)
- `plan-validation-model.md` (conceptual semantic plan validation layer before execution planning)
- `plan-capture-contract.md` (WP-17 enforceable captured-plan contract; docs/tests/tooling-only)
- `plan-capture.schema.json` (WP-17 versioned schema for captured-plan artifacts)

## System Architecture

See the full architecture overview:

`docs/architecture/smartstat-architecture.md`

## Explicit non-scope

- No SmartStat runtime behavior changes
- No resolver/planner execution changes
- No runtime integration dependencies

This layer is documentation and extraction tooling only.
