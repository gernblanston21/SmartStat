# OnAir Filter Grammar Dictionary

## Purpose / Overview

Deterministic filter/qualifier grammar reference from QUALIFIER_TO_FILTER sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: QUALIFIER_TO_FILTER

## Extraction notes / normalization notes

- Header normalization: `Filter (Qualifier) Syntax` and `Qualifier/Filter` are normalized to `filter_name`.
- `parameter_type` uses a controlled vocabulary: `none`, `enum`, `integer`, `integer_or_range`, `year`, `date`, `comparison`, `score_margin`, `team_tricode`, `entity_reference`, `relative_period`, `mixed`, `unknown`.
- Parameter typing in this cleanup pass was refined only where workbook parameters/examples provided clear evidence; ambiguous cases remain `unknown`.

## Extracted reference

| filter_name | parameter_type | parameter_examples | league | notes |
| --- | --- | --- | --- | --- |
| active | none | active | MLB |  |
| after_inning | enum | lead, trail, tie \| after_inning(lead) | MLB |  |
| all_star_break | enum | before, after \| all_star_break(after) | MLB |  |
| all_time | mixed | (REG, [year]), (PST, [year]) 'year = cutoff for stats eg. 2004' \| all_time, all_time(REG, 2002), all_time(2009) | MLB |  |
| asLHB | none | asLHB, asRHB | MLB |  |
| asRHB | none | asLHB, asRHB | MLB |  |
| as_pos | enum | c, 1b, 2b, ss, 3b, lf, rf, cf, dh, ph, pr, p \| as_pos(1b), as_pos(ph) | MLB |  |
| batted_direction3 | enum | left, center, right \| batted_direction3(left) | MLB |  |
| career | enum | reg, post \| career, career(reg) | MLB |  |
| conference | none | conference | MLB |  |
| count | enum | 0-0, 1-0, 2-0, 3-0, 0-1, 0-2, 1-1, 1-2, 2-1, 3-1, 3-2, 2-2 \| count(0-0), count(0-2) | MLB |  |
| count_advantage | enum | even, hitter, pitcher \| count_advantage(even) | MLB |  |
| event_margin | score_margin | <#, >#, =>#, <=#, =#, and #-#, !#-# \| event_margin(<2), event_margin(=3) | MLB |  |
| final_margin | score_margin | <#, >#, =>#, <=#, =#, and #-#, !#-# \| final_margin(<3), final_margin(=5) | MLB |  |
| game | none | game | MLB |  |
| game_extra_innings | enum | yes, no \| game_extra_innings(yes) | MLB |  |
| game_start | enum | day, night \| game_start(day) | MLB |  |
| innings | integer_or_range | # \| innings(3), innings(7-25) | MLB |  |
| is_qualified_batting_al | none | is_qualified_batting_al | MLB |  |
| is_qualified_batting_nl | none | is_qualified_batting_nl | MLB |  |
| is_qualified_hitting | none | is_qualified_hitting | MLB |  |
| is_qualified_pitching | none | is_qualified_pitching | MLB |  |
| is_qualified_pitching_al | none | is_qualified_pitching_al | MLB |  |
| is_qualified_pitching_nl | none | is_qualified_pitching_nl | MLB |  |
| last_game | integer | # \| last_game(1), last_game(5) | MLB |  |
| location | enum | home, away \| location(away) | MLB |  |
| month | mixed | last#, prev#, [month name] 'month name eg. august' \| month(last2), month(july) | MLB |  |
| on_team | team_tricode | TRICODE 'eg. DET' \| on_team(DET) | MLB |  |
| outs | integer | 0, 1, 2 \| outs(2) | MLB |  |
| pa_result | enum | 1b, 2b, 3b, hr, k, bb, ibb, hbp, sf, sb, fc, error, interference, other \| pa_result(2b), pa_result(hr) | MLB |  |
| pitch_category | enum | fastball, offspeed, breaking \| pitch_category(fastball) | MLB |  |
| pitch_location | enum | top-left, top-mid, top-right, mid-left, middle, mid-right, low-left, low-mid, low-right, upper-left, upper-right, lower-left, lower-right \| pitch_location(top-left) | MLB |  |
| pitch_number | comparison | <#, >#, =>#, <=#, =#, and #-#, !#-# \| pitch_number(<4), pitch_number(=8) | MLB |  |
| pitch_type | enum | 4-seam, knuckle_curve, cutter, slider, changeup, 2-seam, slurve, curveball, slow_curve, sweeper, knuckle_ball, forkball, splitter, screwball, eephus \| pitch_type(slider) | MLB |  |
| pos_group | enum | OF, IF, DH, C, P \| pos_group(OF) | MLB |  |
| pos_primary | enum | C, 1B, 2B, SS, 3B, LF, RF, CF, OF, P, SP, RP, IF, DH \| pos_primary(SS) | MLB |  |
| postseason | mixed | last#, prev#, [year] 'year eg. 2022' \| postseason(last2), postseason(2023) | MLB |  |
| risp | enum | yes, no \| risp(yes) | MLB |  |
| rookie | none | rookie | MLB |  |
| runners | enum | empty, 1b, 2b, 3b, 1b-2b, 1b-3b, 2b-3b, loaded \| runners(1b), runners(1b-3b), runners(loaded) | MLB |  |
| runners_on | enum | yes, no \| runners_on(no) | MLB |  |
| runs_allowed | score_margin | <#, >#, =>#, <=#, =#, and #-#, !#-# \| runs_allowed(<3), runs_allowed(=5) | MLB |  |
| runs_scored | score_margin | <#, >#, =>#, <=#, =#, and #-#, !#-# \| runs_scored(>=4) | MLB |  |
| season | mixed | last#, prev#, [year] 'year eg. 2022' \| season, season(last3), season(2024) | MLB |  |
| strikes | integer | 0, 1, 2 \| strikes(2) | MLB |  |
| thru_order | integer | N 'N = times through the order eg. 3' \| thru_order(3) | MLB |  |
| venue | team_tricode | TRICODE 'eg. DET' \| venue(DET) | MLB |  |
| vs | team_tricode | TRICODE 'eg. DET' \| vs(DET) | MLB |  |
| vsLHB | none | vsLHB, vsRHB | MLB |  |
| vsLHP | none | vsLHP, vsRHP | MLB |  |
| vsRHB | none | vsLHB, vsRHB | MLB |  |
| vsRHP | none | vsLHP, vsRHP | MLB |  |
| vs_league | enum | own, inter \| vs_league(own) | MLB |  |
| active | none | active | NHL |  |
| all_star_break | enum | after \| all_star_break(after) | NHL |  |
| all_star_break | enum | before \| all_star_break(before) | NHL |  |
| all_time | mixed | (REG, [year]) or (PST, [year]) year = cutoff for stats \| all_time, all_time(REG, 2002), all_time(2009) | NHL |  |
| assist_per_game_qualifier | none | assist_per_game_qualifier | NHL |  |
| career | enum | reg, post \| career, career(reg) | NHL |  |
| conceded_first | none | conceded_first | NHL |  |
| days_since_game | comparison | <#, >#, =>#, <=#, =#, and #-# \| days_since_game(>3), days_since_game(=3) | NHL |  |
| final_margin | score_margin | <#, >#, =>#, <=#, =#, and #-# \| final_margin(<3), final_margin(=5) | NHL |  |
| game_clock | comparison | <#:##, >#:##, =>#:##, <=#:##, =, #:##and ##:##-#:## \| game_clock(<2:00), game_clock(>=3:00) | NHL |  |
| game_type | enum | regular, playoff \| game_type(playoff) | NHL |  |
| goal_location | enum | LR, LL, LM, UM, UL, UR \| goal_location(LL), goal_location(UR) | NHL |  |
| goals_against_qualifier | none | goals_against_qualifier | NHL |  |
| goals_per_game_qualifier | none | goals_per_game_qualifier | NHL |  |
| last_game | integer | # \| last_game(1), last_game(5) | NHL |  |
| location | enum | away \| location(away) | NHL |  |
| location | enum | home \| location(home) | NHL |  |
| month | mixed | last#, prev#, [month name] \| month(last2), month(july) | NHL |  |
| on_team | team_tricode | TRICODE \| on_team(DET) | NHL |  |
| overtime_games | none | overtime_games | NHL |  |
| penalty_severity | enum | minor, double_minor, major, misconduct, match \| penalty_severity(minor), penalty_severity(double_minor) | NHL |  |
| penalty_type | enum | fighting, charging, boarding, bench, check_to_head, clipping, closing_hand_on_puck, cross_checking, delay_of_game, elbowing, faceoff_violation, high_sticking, holding_the_stick, hooking, instigating, interference, illegal_equipment, kneeing, misconduct, roughing, slashing, too_many_men, tripping, unsportsmanlike_conduct \| penalty_type(cross_checking), penalty_type(hooking) | NHL |  |
| period | enum | 1, 2, 3, OT1, OT2, OT3, OT4, OT5, SO \| period(2), period(OT1) | NHL |  |
| points_earned_qualifier | none | points_earned_qualifier | NHL |  |
| points_per_game_qualifier | none | points_per_game_qualifier | NHL |  |
| pos_group | enum | F, D, G \| pos_group(F), pos_group(G) | NHL |  |
| pos_primary | enum | C, RW, LW, D, G \| pos_primary(RW), pos_primary(D) | NHL |  |
| postseason | mixed | last#, prev#, [year] \| postseason(last2), postseason(2023) | NHL |  |
| rookie | none | rookie | NHL |  |
| save_percentage_qualifier | none | save_percentage_qualifier | NHL |  |
| scored_first | none | scored_first | NHL |  |
| season | mixed | last#, prev#, [year] \| season, season(last3), season(2024) | NHL |  |
| shooting_percentage_qualifier | none | shooting_percentage_qualifier | NHL |  |
| special_teams | enum | PP, EVEN, SH \| special_teams(PP), special_teams(EVEN) | NHL |  |
| team_game | comparison | <#, >#, =>#, <=#, =#, and #-# \| team_game(>30), team_game(<=25) | NHL |  |
| thru_1 | enum | lead, tie, trail \| thru_1(lead), thru_1(tie) | NHL |  |
| thru_2 | enum | lead, tie, trail \| thru_2(lead), thru_2(tie) | NHL |  |
| venue | team_tricode | TRICODE \| venue(DET) | NHL |  |
| vs | team_tricode | TRICODE \| vs(DET) | NHL |  |

## SmartStat Relevance

- Defines deterministic filter grammar tokens and parameter shapes.
- Supports explainable qualifier parsing and normalization.
- Provides source-grounded planner grammar filter slots.
