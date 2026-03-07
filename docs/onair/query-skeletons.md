# OnAir Query Skeleton Grammar

## Purpose / Overview

Canonical query-template shapes extracted from `Available Queries` sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: Available Queries (10), Available Queries (10)

## Extraction notes / normalization notes

- Only rows containing `{{...}}` templates are included.
- Shifted source rows (query text under `Type`) are normalized and documented.
- `MLB_OnAir_v3_Stat_Syntax_v2.xlsx` had 8 shifted query row(s) in `Available Queries`.
- `NHL_OnAir_v3_Stat_Syntax_v2.xlsx` had 8 shifted query row(s) in `Available Queries`.

## Extracted reference

| query_family | skeleton | example | league | source_sheet | notes |
| --- | --- | --- | --- | --- | --- |
| CALENDAR | {{calendar(#,day).entity.filter.measure}} |  | MLB | Available Queries |  |
| CONDITIONAL | {{ %if [query] = [query or value] %then [query or value] %else [query or value] %endif }} |  | MLB | Available Queries |  |
| CUSTOM | {{custom.user-dictionary-variable}} |  | MLB | Available Queries |  |
| GAME HIGH | {{game_high.entity.filter.measure}} |  | MLB | Available Queries |  |
| GAMES WITH | {{games_with.entity.filter.measure}} |  | MLB | Available Queries |  |
| INFO | {{info.entity.attribute}} |  | MLB | Available Queries |  |
| LEADER | {{leader(#,measure).entity.filter.attribute_or_stat}} |  | MLB | Available Queries |  |
| MATH | {{query.entity.filter.measure 'operator'# 'formatter'}} |  | MLB | Available Queries |  |
| PREVIOUS | {{previous.entity.filter.measure}} |  | MLB | Available Queries |  |
| RANK | {{rank.entity.filter.measure}} |  | MLB | Available Queries |  |
| STATS | {{stats.entity.filter.measure}} |  | MLB | Available Queries |  |
| STREAK | {{streak.entity.filter.measure}} |  | MLB | Available Queries |  |
| UNKNOWN | {{ calendar(1, sun).team.game_vs }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(1, sun).time.day_of_month }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(2, mon).team.opp_alias }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(2, tue).team.opp_name }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(3, thu).team.runs }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(3, wed).team.game_result }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(4, fri).team.runs_allowed }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(5, sat).time.day_of_month }} |  | MLB | Available Queries | Query value was shifted under `Type` column in source row. |
| CALENDAR | {{calendar(#,day).entity.filter.measure}} |  | NHL | Available Queries |  |
| CONDITIONAL | {{ %if [query] = [query or value] %then [query or value] %else [query or value] %endif }} |  | NHL | Available Queries |  |
| CUSTOM | {{custom.user-dictionary-variable}} |  | NHL | Available Queries |  |
| GAME HIGH | {{game_high.entity.filter.measure}} |  | NHL | Available Queries |  |
| GAMES WITH | {{games_with.entity.filter.measure}} |  | NHL | Available Queries |  |
| INFO | {{info.entity.attribute}} |  | NHL | Available Queries |  |
| LEADER | {{leader(#,measure).entity.filter.attribute_or_stat}} |  | NHL | Available Queries |  |
| MATH | {{query.entity.filter.measure 'operator'# 'formatter'}} |  | NHL | Available Queries |  |
| PREVIOUS | {{previous.entity.filter.measure}} |  | NHL | Available Queries |  |
| RANK | {{rank.entity.filter.measure}} |  | NHL | Available Queries |  |
| STATS | {{stats.entity.filter.measure}} |  | NHL | Available Queries |  |
| STREAK | {{streak.entity.filter.measure}} |  | NHL | Available Queries |  |
| UNKNOWN | {{ calendar(1, sun).team.game_vs }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(1, sun).time.day_of_month }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(2, mon).team.opp_alias }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(2, tue).team.opp_name }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(3, thu).team.runs }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(3, wed).team.game_result }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(4, fri).team.runs_allowed }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |
| UNKNOWN | {{ calendar(5, sat).time.day_of_month }} |  | NHL | Available Queries | Query value was shifted under `Type` column in source row. |

## SmartStat Relevance

- Supports semantic query-shape dictionaries.
- Supports candidate-resolution context by query family.
- Supports planner grammar scaffolding without planner/runtime execution.
