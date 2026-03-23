# MLB OnAir Profile Extraction Report (PASS_02 FULL INDEX)

## Source

- Source profile: `docs/onair/source/MLB_Network_26.json`
- Pass: `ONAIR_PROFILE_EXTRACTION_MLB_PASS_02_FULL_INDEX`
- Mode: Read-Only Tooling / Semantic-Ingestion
- Runtime impact: None

## PASS_02 Coverage

- Full stat defaults extraction: 771 rows (no sampling)
- Full user dictionary extraction: 187 rows (no sampling)
- Full playbook metadata index: 1862 rows (no sampling)
- Expanded deterministic syntax corpus: full normalized coverage over all template-bearing tabs
  - total template-bearing tabs: 33130
  - unique raw templates: 13433
  - normalized patterns: 437
- Full game plan filters extraction: 28 rows (no sampling)

## Findings

### Measure Catalog Universe

PASS_02 captures the complete MLB profile measure universe with key-pair anchors (`queryEngineMeasureKey` + `onAirCatalogMeasureKey`), descriptions, and deterministic format metadata.

### Operator Phrase Universe

PASS_02 captures the complete user dictionary (`userVariables`) with group context for phrase and shorthand normalization analysis.

### Playbook Index Universe

PASS_02 captures all playbook rows with deterministic metadata and runtime-safe semantic flags (custom placeholders, hybrid text+token content, multi-token composition).

### Syntax Composition Universe

PASS_02 captures full template-tab coverage through a deterministic normalized corpus (not a small sample), while preserving canonical raw template examples and classification metadata.

### Filter Grammar Universe

PASS_02 captures all `gamePlanFilters` rows for player/team time/split grammar references.

## SmartStat Value

Immediately useful as reference/training inputs:
- full measure-key universe
- full operator phrase universe
- full playbook metadata universe
- full syntax composition universe (normalized full coverage)
- full filter grammar universe

Still NOT authorized:
- runtime replacement
- direct INI replacement
- runtime authority
- automatic SmartStat integration

Future enrichment targets (reference only, separate authorized passes):
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_TemplateConfig.ini`
- future supplemental OnAir-derived artifacts

## Boundary Statement

These artifacts are tooling/reference outputs only and do not alter SmartStat runtime behavior or governance.
