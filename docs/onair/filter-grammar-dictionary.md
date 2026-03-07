# OnAir Filter Grammar Dictionary

## Purpose / Overview

Reference list of filter forms and parameter signatures from QUALIFIER_TO_FILTER sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: QUALIFIER_TO_FILTER (3), QUALIFIER_TO_FILTER (2)

## Extraction notes / normalization notes

- Header normalization: `Filter` and `Qualifier/Filter` -> `filter_name`.
- Parameter type classification is conservative and deterministic.

## Extracted reference

| filter_name | parameter_type | parameter_examples | league | source_sheet | notes |
| --- | --- | --- | --- | --- | --- |
| active | none |  | MLB | QUALIFIER_TO_FILTER |  |
| after_inning | enum | lead, trail, tie | MLB | QUALIFIER_TO_FILTER |  |
| all_star_break | enum | before, after | MLB | QUALIFIER_TO_FILTER |  |
| all_time | year | (REG, [year]), (PST, [year]) 'year = cutoff for stats eg. 2004' | MLB | QUALIFIER_TO_FILTER |  |
| asLHB, asRHB | none |  | MLB | QUALIFIER_TO_FILTER |  |
| as_pos | enum | c, 1b, 2b, ss, 3b, lf, rf, cf, dh, ph, pr, p | MLB | QUALIFIER_TO_FILTER |  |
| batted_direction3 | enum | left, center, right | MLB | QUALIFIER_TO_FILTER |  |
| career | enum | reg, post | MLB | QUALIFIER_TO_FILTER |  |
| conference | none |  | MLB | QUALIFIER_TO_FILTER |  |
| count | enum | 0-0, 1-0, 2-0, 3-0, 0-1, 0-2, 1-1, 1-2, 2-1, 3-1, 3-2, 2-2 | MLB | QUALIFIER_TO_FILTER |  |
| count_advantage | enum | even, hitter, pitcher | MLB | QUALIFIER_TO_FILTER |  |
| event_margin | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# | MLB | QUALIFIER_TO_FILTER |  |
| final_margin | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# | MLB | QUALIFIER_TO_FILTER |  |
| game | none |  | MLB | QUALIFIER_TO_FILTER |  |
| game_extra_innings | enum | yes, no | MLB | QUALIFIER_TO_FILTER |  |
| game_start | enum | day, night | MLB | QUALIFIER_TO_FILTER |  |
| innings | unknown | # | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_batting_al | none |  | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_batting_nl | none |  | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_hitting | none |  | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_pitching | none |  | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_pitching_al | none |  | MLB | QUALIFIER_TO_FILTER |  |
| is_qualified_pitching_nl | none |  | MLB | QUALIFIER_TO_FILTER |  |
| last_game | unknown | # | MLB | QUALIFIER_TO_FILTER |  |
| location | enum | home, away | MLB | QUALIFIER_TO_FILTER |  |
| month | enum | last#, prev#, [month name] 'month name eg. august' | MLB | QUALIFIER_TO_FILTER |  |
| on_team | team_tricode | TRICODE 'eg. DET' | MLB | QUALIFIER_TO_FILTER |  |
| outs | enum | 0, 1, 2 | MLB | QUALIFIER_TO_FILTER |  |
| pa_result | enum | 1b, 2b, 3b, hr, k, bb, ibb, hbp, sf, sb, fc, error, interference, other | MLB | QUALIFIER_TO_FILTER |  |
| pitch_category | enum | fastball, offspeed, breaking | MLB | QUALIFIER_TO_FILTER |  |
| pitch_location | enum | top-left, top-mid, top-right, mid-left, middle, mid-right, low-left, low-mid, low-right, upper-left, upper-right, lower-left, lower-right | MLB | QUALIFIER_TO_FILTER |  |
| pitch_number | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# | MLB | QUALIFIER_TO_FILTER |  |
| pitch_type | enum | 4-seam, knuckle_curve, cutter, slider, changeup, 2-seam, slurve, curveball, slow_curve, sweeper, knuckle_ball, forkball, splitter, screwball, eephus | MLB | QUALIFIER_TO_FILTER |  |
| pos_group | enum | OF, IF, DH, C, P | MLB | QUALIFIER_TO_FILTER |  |
| pos_primary | enum | C, 1B, 2B, SS, 3B, LF, RF, CF, OF, P, SP, RP, IF, DH | MLB | QUALIFIER_TO_FILTER |  |
| postseason | year | last#, prev#, [year] 'year eg. 2022' | MLB | QUALIFIER_TO_FILTER |  |
| risp | enum | yes, no | MLB | QUALIFIER_TO_FILTER |  |
| rookie | none |  | MLB | QUALIFIER_TO_FILTER |  |
| runners | enum | empty, 1b, 2b, 3b, 1b-2b, 1b-3b, 2b-3b, loaded | MLB | QUALIFIER_TO_FILTER |  |
| runners_on | enum | yes, no | MLB | QUALIFIER_TO_FILTER |  |
| runs_allowed | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# | MLB | QUALIFIER_TO_FILTER |  |
| runs_scored | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# | MLB | QUALIFIER_TO_FILTER |  |
| season | year | last#, prev#, [year] 'year eg. 2022' | MLB | QUALIFIER_TO_FILTER |  |
| strikes | enum | 0, 1, 2 | MLB | QUALIFIER_TO_FILTER |  |
| thru_order | range | N 'N = times through the order eg. 3' | MLB | QUALIFIER_TO_FILTER |  |
| venue | team_tricode | TRICODE 'eg. DET' | MLB | QUALIFIER_TO_FILTER |  |
| vs | team_tricode | TRICODE 'eg. DET' | MLB | QUALIFIER_TO_FILTER |  |
| vsLHB, vsRHB | none |  | MLB | QUALIFIER_TO_FILTER |  |
| vsLHP, vsRHP | none |  | MLB | QUALIFIER_TO_FILTER |  |
| vs_league | enum | own, inter | MLB | QUALIFIER_TO_FILTER |  |
| active | none |  | NHL | QUALIFIER_TO_FILTER |  |
| all_star_break(after) | enum | after | NHL | QUALIFIER_TO_FILTER | Alias column present. |
| all_star_break(before) | enum | before | NHL | QUALIFIER_TO_FILTER | Alias column present. |
| all_time | year | (REG, [year]) or (PST, [year]) year = cutoff for stats | NHL | QUALIFIER_TO_FILTER |  |
| assist_per_game_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| career | enum | reg, post | NHL | QUALIFIER_TO_FILTER |  |
| conceded_first | none |  | NHL | QUALIFIER_TO_FILTER |  |
| days_since_game | comparison | <#, >#, =>#, <=#, =#, and #-# | NHL | QUALIFIER_TO_FILTER |  |
| final_margin | comparison | <#, >#, =>#, <=#, =#, and #-# | NHL | QUALIFIER_TO_FILTER |  |
| game_clock | comparison | <#:##, >#:##, =>#:##, <=#:##, =, #:##and ##:##-#:## | NHL | QUALIFIER_TO_FILTER |  |
| game_type | enum | regular, playoff | NHL | QUALIFIER_TO_FILTER |  |
| goal_location | enum | LR, LL, LM, UM, UL, UR | NHL | QUALIFIER_TO_FILTER |  |
| goals_against_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| goals_per_game_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| last_game | unknown | # | NHL | QUALIFIER_TO_FILTER |  |
| location(away) | unknown | away | NHL | QUALIFIER_TO_FILTER |  |
| location(home) | unknown | home | NHL | QUALIFIER_TO_FILTER |  |
| month | enum | last#, prev#, [month name] | NHL | QUALIFIER_TO_FILTER |  |
| on_team | team_tricode | TRICODE | NHL | QUALIFIER_TO_FILTER |  |
| overtime_games | none |  | NHL | QUALIFIER_TO_FILTER |  |
| penalty_severity | enum | minor, double_minor, major, misconduct, match | NHL | QUALIFIER_TO_FILTER |  |
| penalty_type | enum | fighting, charging, boarding, bench, check_to_head, clipping, closing_hand_on_puck, cross_checking, delay_of_game, elbowing, faceoff_violation, high_sticking, holding_the_stick, hooking, instigating, interference, illegal_equipment, kneeing, misconduct, roughing, slashing, too_many_men, tripping, unsportsmanlike_conduct | NHL | QUALIFIER_TO_FILTER |  |
| period | enum | 1, 2, 3, OT1, OT2, OT3, OT4, OT5, SO | NHL | QUALIFIER_TO_FILTER |  |
| points_earned_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| points_per_game_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| pos_group | enum | F, D, G | NHL | QUALIFIER_TO_FILTER |  |
| pos_primary | enum | C, RW, LF, D, G | NHL | QUALIFIER_TO_FILTER |  |
| postseason | year | last#, prev#, [year] | NHL | QUALIFIER_TO_FILTER |  |
| rookie | none |  | NHL | QUALIFIER_TO_FILTER |  |
| save_percentage_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| scored_first | none |  | NHL | QUALIFIER_TO_FILTER |  |
| season | year | last#, prev#, [year] | NHL | QUALIFIER_TO_FILTER |  |
| shooting_percentage_qualifier | none |  | NHL | QUALIFIER_TO_FILTER |  |
| special_teams | enum | PP, EVEN, SH | NHL | QUALIFIER_TO_FILTER |  |
| team_game | comparison | <#, >#, =>#, <=#, =#, and #-# | NHL | QUALIFIER_TO_FILTER |  |
| thru_1 | enum | lead, tie, trail | NHL | QUALIFIER_TO_FILTER |  |
| thru_2 | enum | lead, tie, trail | NHL | QUALIFIER_TO_FILTER |  |
| venue | team_tricode | TRICODE | NHL | QUALIFIER_TO_FILTER |  |
| vs | team_tricode | TRICODE | NHL | QUALIFIER_TO_FILTER |  |

## SmartStat Relevance

- Supports semantic filter dictionaries.
- Supports qualifier parameter explainability.
- Supports future planner grammar inputs without resolver/runtime implementation.
