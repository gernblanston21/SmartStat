# OnAir Extraction Tooling

## Purpose

Deterministic extraction tooling for the Phase 0 OnAir semantic reference layer.

## Source inputs

- `docs/onair/source/MLB_OnAir_v3_Stat_Syntax_v3.xlsx`
- `docs/onair/source/NHL_OnAir_v3_Stat_Syntax_v3.xlsx`

## Script

- `extract_onair_reference.py`

## Usage

```bash
python tools/onair/extract_onair_reference.py
```

## Output targets

- `docs/onair/workbook-inventory.md`
- `docs/onair/README.md`
- `docs/onair/measure-dictionary.md`
- `docs/onair/filter-grammar-dictionary.md`
- `docs/onair/alias-map.qualifiers.md`
- `docs/onair/entity-dictionary.md`
- `docs/onair/attribute-dictionary.md`
- `docs/onair/formatter-dictionary.md`
- `docs/onair/query-skeletons.md`
- `docs/onair/onair_semantic_grammar.md`

## Boundaries

- Read-only semantic reference extraction only
- No runtime SmartStat behavior changes
- No resolver/planner execution logic
