# OnAir Qualifier Phrase Alias Map

## Purpose / Overview

Conservative alias phrase mapping to canonical qualifier/filter forms derived from explicit syntax/parameter evidence.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: QUALIFIER_TO_FILTER

## Extraction notes / normalization notes

- v3 sheets do not include a dedicated alias-column map.
- `mapping_basis=explicit` is reserved for direct workbook alias rows (none observed in this phase).
- `mapping_basis=inferred_from_parameters` is used when mapping is derived from available-parameter forms and examples.

## Extracted reference

| alias_phrase | canonical_filter | canonical_parameter | league | example | mapping_basis |
| --- | --- | --- | --- | --- | --- |
| LAST N GAMES | last_game | N | MLB | last_game(5) | inferred_from_parameters |
| LAST N SEASONS | season | lastN | MLB | season(last3) | inferred_from_parameters |
| POST ALL-STAR BREAK | all_star_break | after | MLB | all_star_break(after) | inferred_from_parameters |
| PRE ALL-STAR BREAK | all_star_break | before | MLB | all_star_break(before) | inferred_from_parameters |
| PREVIOUS N SEASONS | season | prevN | MLB | season(prev2) | inferred_from_parameters |
| SINCE ALL-STAR BREAK | all_star_break | after | MLB | all_star_break(after) | inferred_from_parameters |
| LAST N GAMES | last_game | N | NHL | last_game(5) | inferred_from_parameters |
| LAST N SEASONS | season | lastN | NHL | season(last3) | inferred_from_parameters |
| POST ALL-STAR BREAK | all_star_break | after | NHL | all_star_break(after) | inferred_from_parameters |
| PRE ALL-STAR BREAK | all_star_break | before | NHL | all_star_break(before) | inferred_from_parameters |
| PREVIOUS N SEASONS | season | prevN | NHL | season(prev2) | inferred_from_parameters |
| SINCE ALL-STAR BREAK | all_star_break | after | NHL | all_star_break(after) | inferred_from_parameters |

## SmartStat Relevance

- Seeds deterministic qualifier alias normalization.
- Improves candidate-resolution explainability for human phrase variants.
- Preserves mapping provenance (`explicit` vs `inferred_from_parameters`) for downstream planner work.
