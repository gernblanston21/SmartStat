# MLB Learn-Layer Staging Summary (PASS_05)

## Scope

- Pass: `MLB_LEARN_LAYER_STAGING_PASS_05`
- Mode: Read-Only Analysis / Learn-Layer Staging
- Inputs used: PASS_04 MLB-only rebaseline artifacts from `tests/_scratch/onair-mlb-pass04-mlb-only-rebaseline/`.
- Read-only pass: no `.vbs` changes, no `.ini` changes, no runtime integration.

## Findings

- Bucket counts:
  - `safe_to_stage`: 27
  - `manual_review_required`: 280
  - `composite_defer`: 238
  - `rejected_or_not_yet_useful`: 480

### Strongest safe_to_stage candidates
- `measure_156`: `steals` (alias, confidence=high)
- `measure_277`: `gs` (alias, confidence=high)
- `measure_49`: `caught_stealing` (measure, confidence=high)
- `measure_62`: `singles` (measure, confidence=high)
- `measure_140`: `pickoffs` (measure, confidence=high)
- `measure_214`: `fielding_assists` (measure, confidence=high)
- `measure_216`: `fielding_caught_stealing` (measure, confidence=high)
- `measure_217`: `fielding_double_plays` (measure, confidence=high)
- `measure_226`: `fielding_putouts` (measure, confidence=high)
- `measure_437`: `balks` (measure, confidence=high)
- `measure_449`: `pitcher_blown_saves` (measure, confidence=high)
- `measure_451`: `pitcher_caught_stealing` (measure, confidence=high)

### Strongest manual_review_required candidates
- `measure_2`: `at_bats_bases_empty` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_3`: `at_bats_bases_loaded` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_4`: `at_bats_runners` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_5`: `at_bats_runners_in_scoring_position_two_outs` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_9`: `batter_foul_balls` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_11`: `batter_inzone_contact_percentage` (Context/split dependency or non-high confidence requires manual semantic review.)
- `measure_12`: `batter_inzone_swing_percentage` (Context/split dependency or non-high confidence requires manual semantic review.)
- `measure_13`: `batter_inzone_whiff_percentage` (Context/split dependency or non-high confidence requires manual semantic review.)
- `measure_14`: `batter_outofzone_contact_percentage` (Context/split dependency or non-high confidence requires manual semantic review.)
- `measure_15`: `batter_outofzone_whiff_percentage` (Context/split dependency or non-high confidence requires manual semantic review.)
- `measure_19`: `batter_times_chased` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)
- `measure_22`: `at_bats` (High-confidence missing candidate lacks direct cross-source corroboration or may be context-dependent.)

### Strongest composite_defer candidates
- `user_13`: `RATIO` (operator_phrase_composite)
- `user_111`: `AB PER HR` (operator_phrase_composite)
- `user_112`: `AB PER WALK` (operator_phrase_composite)
- `user_113`: `AB PER STRIKEOUT` (operator_phrase_composite)
- `composite_syntax_2_7`: `{{info.time.season}}: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}` (playbook_syntax_composite)
- `composite_syntax_29_5`: `PITCHES/PA - {{info.time.season}}` (playbook_syntax_composite)
- `composite_syntax_34_6`: `WITH {{info.team.name}}: {{stats.player.season.on_team.batting_average}}, {{stats.player.season.on_team.home_runs}} {{custom.header_hr}}, {{stats.player.season.on_team.rbis}} {{custom.header_rbi}}` (playbook_syntax_composite)
- `composite_syntax_41_7`: `{{stats.player.season.hits}}/{{stats.player.season.at_bats}}` (playbook_syntax_composite)
- `composite_syntax_61_6`: `[REGULAR SEASON: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}` (playbook_syntax_composite)
- `composite_syntax_77_5`: `{{info.time.season}} vs {{info.team(opp).name}} ({{stats.player.season.vs.games_played}} GAMES)` (playbook_syntax_composite)
- `composite_syntax_85_5`: `{{info.time.season}} w/ RISP` (playbook_syntax_composite)
- `composite_syntax_86_5`: `{{info.time.season}} w/ 2 OUTS & RISP` (playbook_syntax_composite)

### Notable exclusions/rejections
- `measure_1`: `rbi_opponent` (Missing candidate is not high confidence in PASS_04 baseline.)
- `measure_6`: `total_distance_average` (Missing candidate is not high confidence in PASS_04 baseline.)
- `measure_7`: `batter_chase_percentage` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_8`: `batter_contact_percentage` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_16`: `batter_swing_percentage` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_17`: `batter_swings` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_18`: `batter_swings_contact` (Missing candidate is not high confidence in PASS_04 baseline.)
- `measure_20`: `batter_whiff_percentage` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_21`: `batter_whiffs` (Missing candidate is not high confidence in PASS_04 baseline.)
- `measure_33`: `batting_average` (Already covered by existing MLB SmartStat mapping universe (direct_match).)
- `measure_34`: `batting_average_bases_empty` (Missing candidate is not high confidence in PASS_04 baseline.)
- `measure_35`: `batting_average_bases_loaded` (Missing candidate is not high confidence in PASS_04 baseline.)

## SmartStat Value

- Produces a review-ready, deterministic staging queue for a future controlled learn-layer proposal pass.
- Separates low-risk staging candidates from ambiguity and composite risk before any INI mutation is considered.
- Safer than direct INI edits because all uncertain and composite candidates are explicitly deferred.

## Governance Boundary

- No runtime changes
- No INI changes
- No automatic integration

## Recommended Next Pass

- `MLB_LEARN_LAYER_CONTROLLED_PROPOSAL_PASS_06`
- Scope: convert only `safe_to_stage` candidates into a bounded, human-reviewable INI proposal patch (still non-runtime).
