# OnAir Extraction Tooling

This directory contains deterministic extraction helpers for the Phase 0 OnAir semantic reference layer.

## Script

- `extract_onair_reference.py`

## What It Does

- Parses MLB/NHL OnAir XLSX files using standard library XML readers
- Inventories actual workbook sheet names/order/header structure
- Extracts reference scaffolding data (measures, filters, aliases, entities, attributes, formatters, queries)
- Writes documentation artifacts under `docs/onair/`
- Writes structured extraction output to `docs/onair/_extracted/onair_reference.extracted.json`

## Scope Guardrail

This is read-only semantic reference extraction tooling.
It does not implement runtime SmartStat behavior, planner execution, or resolver logic.

## Usage

From repo root:

```powershell
python tools/onair/extract_onair_reference.py
```
