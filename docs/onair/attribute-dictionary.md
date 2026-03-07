# OnAir Attribute Dictionary

## Purpose / Overview

Attribute references across player/coach, team, and time sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: Player-Coach Attributes (5), Team Attributes (6), Time Attribute (7), Player-Coach Attributes (5), Team Attributes (6), Time Attribute (7)

## Extraction notes / normalization notes

- Header normalization aligns source naming differences (`Attribute` vs `Measure`).
- `Time Attribute` source headers (`Formatter`/`Function`) are mapped conservatively into attribute rows.

## Extracted reference

| attribute_name | entity_type | league | source_sheet | notes |
| --- | --- | --- | --- | --- |
| abbr_name | player | MLB | Player-Coach Attributes |  |
| age | player | MLB | Player-Coach Attributes |  |
| bat_hand | player | MLB | Player-Coach Attributes |  |
| birthdate | player | MLB | Player-Coach Attributes |  |
| college | player | MLB | Player-Coach Attributes |  |
| draft_pick | player | MLB | Player-Coach Attributes |  |
| draft_round | player | MLB | Player-Coach Attributes |  |
| draft_team_alias | player | MLB | Player-Coach Attributes |  |
| draft_team_name | player | MLB | Player-Coach Attributes |  |
| draft_year | player | MLB | Player-Coach Attributes |  |
| experience | player | MLB | Player-Coach Attributes |  |
| height | player | MLB | Player-Coach Attributes |  |
| high_school | player | MLB | Player-Coach Attributes |  |
| hometown | player | MLB | Player-Coach Attributes |  |
| jersey_number | player | MLB | Player-Coach Attributes |  |
| pitcher_hand | player | MLB | Player-Coach Attributes |  |
| position | player | MLB | Player-Coach Attributes |  |
| primary_position | player | MLB | Player-Coach Attributes |  |
| pro_debut | player | MLB | Player-Coach Attributes |  |
| throw_hand | player | MLB | Player-Coach Attributes |  |
| weight | player | MLB | Player-Coach Attributes |  |
| first_name | player/coach | MLB | Player-Coach Attributes |  |
| full_name | player/coach | MLB | Player-Coach Attributes |  |
| last_name | player/coach | MLB | Player-Coach Attributes |  |
| other_team_alias_list | player/coach | MLB | Player-Coach Attributes |  |
| other_team_list | player/coach | MLB | Player-Coach Attributes |  |
| team_alias_list | player/coach | MLB | Player-Coach Attributes |  |
| team_name_list | player/coach | MLB | Player-Coach Attributes |  |
| alias | team | MLB | Team Attributes |  |
| altcity | team | MLB | Team Attributes |  |
| altname | team | MLB | Team Attributes |  |
| city | team | MLB | Team Attributes |  |
| name | team | MLB | Team Attributes |  |
| venue_city | team | MLB | Team Attributes |  |
| venue_country | team | MLB | Team Attributes |  |
| venue_name | team | MLB | Team Attributes |  |
| venue_state | team | MLB | Team Attributes |  |
| day_of_week | time | MLB | Time Attribute | prev# |
| month | time | MLB | Time Attribute | prev# |
| season | time | MLB | Time Attribute | prev# |
| year | time | MLB | Time Attribute | prev# |
| abbr_name | player | NHL | Player-Coach Attributes |  |
| age | player | NHL | Player-Coach Attributes |  |
| birthdate | player | NHL | Player-Coach Attributes |  |
| college | player | NHL | Player-Coach Attributes |  |
| draft_pick | player | NHL | Player-Coach Attributes |  |
| draft_round | player | NHL | Player-Coach Attributes |  |
| draft_team_alias | player | NHL | Player-Coach Attributes |  |
| draft_team_name | player | NHL | Player-Coach Attributes |  |
| draft_year | player | NHL | Player-Coach Attributes |  |
| experience | player | NHL | Player-Coach Attributes |  |
| handedness | player | NHL | Player-Coach Attributes |  |
| height | player | NHL | Player-Coach Attributes |  |
| high_school | player | NHL | Player-Coach Attributes |  |
| hometown | player | NHL | Player-Coach Attributes |  |
| jersey_number | player | NHL | Player-Coach Attributes |  |
| position | player | NHL | Player-Coach Attributes |  |
| primary_position | player | NHL | Player-Coach Attributes |  |
| rookie_year | player | NHL | Player-Coach Attributes |  |
| weight | player | NHL | Player-Coach Attributes |  |
| first_name | player/coach | NHL | Player-Coach Attributes |  |
| full_name | player/coach | NHL | Player-Coach Attributes |  |
| last_name | player/coach | NHL | Player-Coach Attributes |  |
| other_team_alias_list | player/coach | NHL | Player-Coach Attributes |  |
| other_team_list | player/coach | NHL | Player-Coach Attributes |  |
| team_alias_list | player/coach | NHL | Player-Coach Attributes |  |
| team_name_list | player/coach | NHL | Player-Coach Attributes |  |
| alias | team | NHL | Team Attributes |  |
| altcity | team | NHL | Team Attributes |  |
| altname | team | NHL | Team Attributes |  |
| city | team | NHL | Team Attributes |  |
| name | team | NHL | Team Attributes |  |
| venue_city | team | NHL | Team Attributes |  |
| venue_country | team | NHL | Team Attributes |  |
| venue_name | team | NHL | Team Attributes |  |
| venue_state | team | NHL | Team Attributes |  |
| day_of_week | time | NHL | Time Attribute | prev# |
| month | time | NHL | Time Attribute | prev# |
| season | time | NHL | Time Attribute | prev# |
| year | time | NHL | Time Attribute | prev# |

## SmartStat Relevance

- Supports semantic attribute dictionaries.
- Supports explainable entity-attribute chains.
- Supports planner grammar attribute slots without resolver/runtime behavior.
