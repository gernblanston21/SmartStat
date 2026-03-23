# MLB Learn-Layer Canonical Model Rebaseline Report (PASS_07)

## Scope

- Pass: `MLB_LEARN_LAYER_CANONICAL_MODEL_REBASELINE_PASS_07`
- Inputs: PASS_05 staging artifacts, PASS_04 MLB-only rebaseline artifacts, and SmartStat mapping reference files.
- Mode: read-only proposal correction and structure modeling only.
- No SmartStat files were modified.

## Section Role Findings

- `[CATEGORY_TO_MEASURE]`: canonical section; `KEY=measure_value` for non-pitcher category/stat concepts.
- `[CATEGORY_TO_MEASURE_ALIASES]`: alias section; `ALIAS_KEY=CANONICAL_KEY` (indirection to existing canonical key).
- `[CATEGORY_TO_MEASURE_PITCHER]`: canonical pitcher section; `KEY=measure_value` for pitcher-domain concepts.
- `[CATEGORY_TO_MEASURE_PITCHER_ALIASES]`: alias pitcher section; `ALIAS_KEY=CANONICAL_PITCHER_KEY`.
- `[QUALIFIER_TO_FILTER]`: canonical qualifier section; `QUALIFIER_LABEL=filter_expression`.
- `[QUALIFIER_TO_FILTER_ALIASES]`: alias qualifier section; `ALIAS_LABEL=CANONICAL_QUALIFIER_LABEL`.
- `[QUALIFIER_FILTER_VALUE]`: qualifier display/value templates; `QUALIFIER_KEY=render_template/value_expression`.
- `[PENDING*]` sections in `SmartStat_Mappings.learn.ini`: staging placeholders (currently empty), not semantic destination authority.

Key/value direction model used in this pass:
- canonical sections define semantic destination keys and concrete values
- alias sections reference canonical keys, not measure values and not self-references

## Corrected Classification Findings

- Reclassified safe candidates: **27**
  - `alias_to_existing_canonical`: 1
  - `manual_review_required`: 8
  - `new_canonical_category_measure`: 6
  - `new_canonical_pitcher_category_measure`: 1
  - `reject`: 11
- Survived as proposal-eligible under corrected model: **8**
- Downgraded/rejected/manual under corrected model: **19**
- Require canonical-target decision first: **8**

Surviving proposal-eligible candidates (sample):
- `measure_62` -> `CATEGORY_TO_MEASURE` `SINGLES=singles` (new_canonical_category_measure, high)
- `user_107` -> `CATEGORY_TO_MEASURE_ALIASES` `GO-AHEAD RBI=GO AHEAD RBI` (alias_to_existing_canonical, high)
- `measure_214` -> `CATEGORY_TO_MEASURE` `ASSISTS=fielding_assists` (new_canonical_category_measure, medium)
- `measure_217` -> `CATEGORY_TO_MEASURE` `DOUBLE PLAYS=fielding_double_plays` (new_canonical_category_measure, medium)
- `measure_226` -> `CATEGORY_TO_MEASURE` `PUTOUTS=fielding_putouts` (new_canonical_category_measure, medium)
- `measure_449` -> `CATEGORY_TO_MEASURE_PITCHER` `BLOWN SAVES=pitcher_blown_saves` (new_canonical_pitcher_category_measure, high)
- `measure_762` -> `CATEGORY_TO_MEASURE` `GAME WINNING RBI=game_winning_rbi` (new_canonical_category_measure, high)
- `measure_769` -> `CATEGORY_TO_MEASURE` `GO AHEAD RBI=go_ahead_rbi` (new_canonical_category_measure, high)

Downgraded/manual/reject candidates (sample):
- `user_2` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)
- `user_4` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)
- `user_6` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)
- `user_40` -> `reject` (Duplicate user phrase evidence; first occurrence retained for modeling.)
- `user_41` -> `reject` (Duplicate user phrase evidence; first occurrence retained for modeling.)
- `measure_49` -> `manual_review_required` (Concept collides across runner/fielding/pitcher semantics; canonical destination is not singular.)
- `user_69` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)
- `user_94` -> `manual_review_required` (Dependent measure concept is manual-review gated; alias cannot be safely promoted.)
- `user_95` -> `manual_review_required` (Dependent measure concept is manual-review gated; alias cannot be safely promoted.)
- `user_119` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)
- `user_123` -> `manual_review_required` (User phrase maps to multiple candidate measure concepts; canonical dependency is ambiguous.)
- `user_138` -> `reject` (User phrase equals proposed canonical key; alias entry would be redundant.)

## Structural Error Findings

- `SELF_REFERENTIAL_ALIAS`: `GS=GS` -> Alias sections should map alternate keys to canonical keys; self-referential aliasing adds no indirection and is structurally unsound for mutation planning.
- `SELF_REFERENTIAL_ALIAS`: `STEALS=STEALS` -> Alias must resolve to canonical key (e.g., STEALS=SB); self-reference bypasses canonical model.
- `STAGING_SECTION_AS_SEMANTIC_DESTINATION`: `Treating PENDING_ALIASES as final semantic target without canonical dependency modeling` -> PENDING sections are staging surfaces, not semantic authority; canonical/alias destination sections must be modeled first.

This pass prevents PASS_06-style errors by requiring:
- canonical/alias role resolution before staging section placement
- alias-to-canonical dependencies (no alias self-reference)
- fail-closed downgrade when canonical destination is ambiguous

## Governance Boundary

- No runtime changes
- No INI changes
- No automatic integration

## Recommended Next Pass

- `MLB_LEARN_LAYER_CANONICAL_PATCH_PACKET_PASS_08`
- Scope: build a bounded human-approval patch packet from PASS_07 surviving entries, with explicit canonical dependencies and manual-gate checklist for ambiguous items.
