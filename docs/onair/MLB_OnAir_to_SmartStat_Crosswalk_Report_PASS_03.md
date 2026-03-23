# MLB OnAir to SmartStat Crosswalk Report (PASS_03)

## Scope

- Pass: `ONAIR_PROFILE_TO_SMARTSTAT_CROSSWALK_ANALYSIS_PASS_03`
- Mode: Read-Only Analysis / Semantic Crosswalk
- This pass compares PASS_02 MLB OnAir extraction artifacts against current governed SmartStat semantic/config files.
- This pass is analysis-only and does not modify runtime behavior or any INI governance file.

### Compared Inputs

- OnAir measure source: `docs/onair/MLB_OnAir_Measure_Catalog_Full.json`
- OnAir user dictionary source: `docs/onair/MLB_OnAir_UserDictionary_Full.json`
- OnAir playbook syntax source: `docs/onair/MLB_OnAir_Playbook_Syntax_Corpus_Full.json`
- OnAir filter source: `docs/onair/MLB_OnAir_GamePlanFilters_Full.json`
- SmartStat governed file: `SmartStat_Mappings.ini`
- SmartStat governed file: `SmartStat_Mappings.learn.ini`
- SmartStat governed file: `SmartStat_MappingsNBA.ini`
- SmartStat governed file: `SmartStat_MappingsNBA.learn.ini`
- SmartStat governed file: `SmartStat_MappingsNHL.ini`
- SmartStat governed file: `SmartStat_MappingsNHL.learn.ini`
- SmartStat governed file: `SmartStat_StaticOverrides.ini`
- SmartStat governed file: `SmartStat_TemplateConfig.ini`

## Findings

### Count Summaries

- Measure crosswalk rows: **771**
  - `direct_match`: 77
  - `likely_alias_match`: 74
  - `missing_candidate`: 302
  - `needs_manual_review`: 318
- User dictionary crosswalk rows: **187**
  - `direct_match`: 64
  - `likely_alias_match`: 1
  - `missing_candidate`: 68
  - `needs_manual_review`: 54
- Composite candidate rows: **224**
  - `custom_header_placeholder`: 8
  - `multi_token_composition`: 2
  - `text_token_hybrid`: 214
- Filter/qualifier crosswalk rows: **28**
  - `direct_match`: 16
  - `missing_candidate`: 8
  - `needs_manual_review`: 4

### Strongest Direct Measure Matches (Sample)

- `on_base_percentage` -> `on_base_percentage` (direct_match)
- `on_base_plus_slugging` -> `on_base_plus_slugging` (direct_match)
- `plate_appearances` -> `plate_appearances` (direct_match)
- `pitches_faced` -> `pitches_faced` (direct_match)
- `runs` -> `runs` (direct_match)
- `rbis` -> `rbis` (direct_match)
- `steals_per_game` -> `steals_per_game` (direct_match)
- `batter_swing_percentage` -> `batter_swing_percentage` (direct_match)

### Strongest Measure Alias Opportunities (Sample)

- `on_base_percentage_differential` -> candidate `on_base_percentage` (score=0.75)
- `on_base_percentage_home` -> candidate `on_base_percentage` (score=0.75)
- `on_base_percentage_road` -> candidate `on_base_percentage` (score=0.75)
- `on_base_percentage_lhp` -> candidate `on_base_percentage` (score=0.75)
- `on_base_percentage_rhp` -> candidate `on_base_percentage` (score=0.75)
- `batter_inzone_contact_percentage` -> candidate `batter_contact_percentage` (score=0.75)
- `on_base_plus_slugging_home` -> candidate `on_base_plus_slugging` (score=0.8)
- `on_base_plus_slugging_road` -> candidate `on_base_plus_slugging` (score=0.8)

### Strongest Missing Measure Candidates (Sample)

- `batter_home_run_to_fly_ball_percentage` (composite_measure_candidate)
- `line_drive_outs` (single_measure)
- `plate_apperances_home` (single_measure)
- `plate_apperances_per_walk` (composite_measure_candidate)
- `plate_apperances_per_strikeout` (composite_measure_candidate)
- `plate_apperances_per_rbi` (composite_measure_candidate)
- `plate_apperances_road` (single_measure)
- `hits_doubles_percentage` (single_measure)

### User Dictionary Highlights (Sample)

- Alias candidate: `EXTRA BASE HITS ALLOWED` -> `extra_base_hits` (leaderboard_header_phrase)
- Missing phrase candidate: `ERRORS` (leaderboard_header_phrase)
- Missing phrase candidate: `IR SCORED PCT` (leaderboard_header_phrase)
- Missing phrase candidate: `PITCHING WAR` (leaderboard_header_phrase)
- Missing phrase candidate: `GROUNDED INTO DP` (leaderboard_header_phrase)

### Composite Syntax Candidate Highlights (Sample)

- `custom_header_placeholder`: `{{info.team.alias}}`
- `text_token_hybrid`: `{{info.time.season}} ({{stats.player.season.games_played}} GAMES)`
- `custom_header_placeholder`: `{{info.time.season}}: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}`
- `text_token_hybrid`: `CAREER ({{info.player.experience | ordinal}} SEASON)`
- `text_token_hybrid`: `{{info.time.month}} ({{stats.player.season.month.at_bats}} AB)`
- `text_token_hybrid`: `1st HALF ({{stats.player.season.all_star_break(before).games_played}} GAMES)`
- `text_token_hybrid`: `2nd HALF ({{stats.player.season.all_star_break(after).games_played}} GAMES)`
- `text_token_hybrid`: `AT HOME {{info.time.season}}`
- `text_token_hybrid`: `ON THE ROAD {{info.time.season}}`
- `text_token_hybrid`: `DAY GAMES ({{stats.player.season.game_start(day).games_played}})`

### Filter / Qualifier Alignment Highlights (Sample)

- Direct: `Season` / `season`
- Direct: `Season` / `season`
- Direct: `Home` / `location(home)`
- Direct: `Last Season` / `season(prev1)`
- Direct: `Home` / `location(home)`
- Direct: `Last Season` / `season(prev1)`
- Direct: `Away` / `location(away)`
- Direct: `Postseason` / `postseason`

## SmartStat Value

Immediately enabled by PASS_03 artifacts:
- deterministic shortlist generation for `SmartStat_Mappings.learn.ini` candidate staging
- deterministic split between direct reuse vs alias candidate vs missing candidate
- deterministic identification of composite syntax candidates that exceed one-to-one stat mapping
- deterministic filter/qualifier vocabulary alignment map for later bounded normalization passes

Still requires future explicitly authorized passes:
- any direct INI mutation (`SmartStat_Mappings.ini`, `SmartStat_Mappings.learn.ini`, `SmartStat_TemplateConfig.ini`)
- any runtime integration or behavioral change
- any automatic mapping insertion

## Governance Boundary

- No runtime changes made
- No SmartStat `.vbs` changes made
- No INI changes made
- This pass does not authorize integration; it only produces crosswalk analysis artifacts

## Recommended Next Pass

- `MLB_ONAIR_TO_SMARTSTAT_LEARN_LAYER_CANDIDATE_STAGING_PASS_04`
- Scope recommendation: stage only high-confidence `likely_alias_match` and `missing_candidate` rows into a governance-reviewed learn-layer candidate package (analysis-first, no runtime changes).
