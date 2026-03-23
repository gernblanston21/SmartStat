# MLB OnAir Profile Extraction Report (PASS_01)

## Source

- Source profile: `docs/onair/source/MLB_Network_26.json`
- Mode: Read-Only Tooling / Semantic-Ingestion
- Runtime impact: None (no runtime/VBScript behavior changes)

## Extracted Sections

- `settings.statDefaults`: 771 rows
- `settings.userDictionary.userVariables`: 187 rows
- `playbook`: 1862 rows
- `settings.gamePlanFilters`: 28 filter rows (player/team time/split buckets)

## Findings

### A. Measure Catalog

- Deterministic key anchors extracted: `queryEngineMeasureKey`, `onAirCatalogMeasureKey`.
- Formatting metadata preserved (`formatType`, decimal/suffix, live/sort flags).
- PASS_01 output is bounded and deterministic for safe semantic-ingestion.

### B. Operator Phrase Dictionary

- Operator-facing user variables extracted with grouping context.
- Useful as alias/shorthand candidates for SmartStat semantic references.

### C. Composite Syntax Corpus

- Real playbook syntax examples extracted for single token, multi token, text+token hybrid, and custom header placeholders.
- Base page/write pattern metadata preserved in a bounded playbook index extract.

### D. Filter Grammar Candidates

- `gamePlanFilters` name/filter/type rows extracted across player/team time/split buckets.
- Suitable for qualifier/filter normalization candidate references.

## SmartStat Value

Immediately useful:
- measure key crosswalk seeds for mapping work
- operator phrase dictionary seeds for alias capture
- playbook syntax corpus seeds for template token understanding
- filter grammar candidate seeds for normalization references

Boundaries:
- tooling/reference only
- not runtime replacement
- not INI governance replacement

Potential future enrichment targets (separate authorized passes):
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_TemplateConfig.ini`
- supplemental OnAir-derived artifacts

## Governance Note

The MLB OnAir export is valuable as a semantic/training/reference source, but it is not yet a direct runtime replacement for current INI governance.
