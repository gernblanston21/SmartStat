# OnAir Formatter Dictionary

## Purpose / Overview

Formatter token reference extracted from Formatters sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: Formatters

## Extraction notes / normalization notes

- `usage_context` classification: `pipe_formatter`, `measure_scope_modifier`, or `formatter`.

## Extracted reference

| formatter_name | description | usage_context | league |
| --- | --- | --- | --- |
| (ingame) | Add after a measure to include stats from the current game | measure_scope_modifier | MLB |
| (pregame) | Add after a measure to exclude current game stats | measure_scope_modifier | MLB |
| *N | Multiply N to Measure | formatter | MLB |
| +N | Add N to Measure | formatter | MLB |
| -N | Subtract N to Measure | formatter | MLB |
| /N | Divide N to Measure | formatter | MLB |
| \| day_long | applies to any "day_of_week" attribute | pipe_formatter | MLB |
| \| day_short | applies to any "day_of_week" attribute | pipe_formatter | MLB |
| \| file_path | Replaces spaces with _ and removes any special characters | pipe_formatter | MLB |
| \| flex_long | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | MLB |
| \| flex_short | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | MLB |
| \| long_noyear | applies to any "date" attribute | pipe_formatter | MLB |
| \| long_year | applies to any "date" attribute | pipe_formatter | MLB |
| \| lowercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB |
| \| ordinal | Adds proper ordinal to number | pipe_formatter | MLB |
| \| short_noyear | applies to any "date" attribute | pipe_formatter | MLB |
| \| short_year | applies to any "date" attribute | pipe_formatter | MLB |
| \| smallcaps | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB |
| \| title | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB |
| \| uppercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB |
| *N | Multiply N to Measure | formatter | NHL |
| +N | Add N to Measure | formatter | NHL |
| -N | Subtract N to Measure | formatter | NHL |
| /N | Divide N to Measure | formatter | NHL |
| \| day_long | applies to any "day_of_week" attribute | pipe_formatter | NHL |
| \| day_short | applies to any "day_of_week" attribute | pipe_formatter | NHL |
| \| file_path | Replaces spaces with _ and removes any special characters | pipe_formatter | NHL |
| \| flex_long | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | NHL |
| \| flex_short | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | NHL |
| \| long_noyear | applies to any "date" attribute | pipe_formatter | NHL |
| \| long_year | applies to any "date" attribute | pipe_formatter | NHL |
| \| lowercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL |
| \| ordinal | Adds proper ordinal to number | pipe_formatter | NHL |
| \| short_noyear | applies to any "date" attribute | pipe_formatter | NHL |
| \| short_year | applies to any "date" attribute | pipe_formatter | NHL |
| \| smallcaps | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL |
| \| title | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL |
| \| uppercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL |

## SmartStat Relevance

- Defines formatter vocabulary and usage shape.
- Supports explainable expression post-processing semantics.
- Provides planner grammar formatter references without runtime coupling.
