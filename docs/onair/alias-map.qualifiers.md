# OnAir Qualifier Phrase Alias Map

## Purpose / Overview

Conservative alias phrase mapping to canonical qualifier/filter forms derived from explicit syntax/parameter evidence.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: QUALIFIER_TO_FILTER

## Extraction notes / normalization notes

- v3 sheets do not provide a dedicated alias column; mappings here are marked as inferred when derived from explicit parameter semantics.
- No free-form alias invention is performed beyond strongly indicated parameterized forms.

## Extracted reference

| alias_phrase | canonical_filter | canonical_parameter | league | example | notes |
| --- | --- | --- | --- | --- | --- |
| LAST N GAMES | last_game | N | MLB | last_game(1), last_game(5) | Inferred from numeric placeholder examples. |
| LAST N SEASONS | season | lastN | MLB | season, season(last3), season(2024) | Inferred from `last#` parameter pattern. |
| POST ALL-STAR BREAK | all_star_break | after | MLB | all_star_break(after) | Inferred from explicit before/after parameter semantics. |
| PRE ALL-STAR BREAK | all_star_break | before | MLB | all_star_break(after) | Inferred from explicit before/after parameter semantics. |
| PREVIOUS N SEASONS | season | prevN | MLB | season, season(last3), season(2024) | Inferred from `prev#` parameter pattern. |
| SINCE ALL-STAR BREAK | all_star_break | after | MLB | all_star_break(after) | Inferred from explicit before/after parameter semantics. |
| LAST N GAMES | last_game | N | NHL | last_game(1), last_game(5) | Inferred from numeric placeholder examples. |
| LAST N SEASONS | season | lastN | NHL | season, season(last3), season(2024) | Inferred from `last#` parameter pattern. |
| POST ALL-STAR BREAK | all_star_break | after | NHL | all_star_break(after) | Inferred from explicit before/after parameter semantics. |
| PRE ALL-STAR BREAK | all_star_break | before | NHL | all_star_break(before) | Inferred from explicit before/after parameter semantics. |
| PREVIOUS N SEASONS | season | prevN | NHL | season, season(last3), season(2024) | Inferred from `prev#` parameter pattern. |
| SINCE ALL-STAR BREAK | all_star_break | after | NHL | all_star_break(after) | Inferred from explicit before/after parameter semantics. |

## SmartStat Relevance

- Seeds deterministic qualifier alias normalization.
- Improves candidate-resolution explainability for human phrase variants.
- Provides a controlled synonym layer for future semantic/planner work.
