# MLB OnAir Semantic Lane Definition Report (PASS_02)

## Scope

- Pass: `MLB_ONAIR_SEMANTIC_LAYER_LANE_DEFINITION_PASS_02`
- Mode: read-only lane-definition/analysis only.
- Primary evidence source: PASS_01 classification artifacts.
- Supporting evidence: PASS_09 patch signoff summary and PASS_12 runtime post-promotion validation summary.
- No runtime changes, no INI changes, no resolver/composite implementation changes were made.

## PASS_01 Classification Recap

- Total candidates: **4583**
- `runtime_mapping_candidate`: 12 (0.26%)
- `alias_candidate`: 3 (0.07%)
- `composite_candidate`: 633 (13.81%)
- `semantic_layer_candidate`: 1890 (41.24%)
- `tooling_only`: 499 (10.89%)
- `ambiguous`: 1546 (33.73%)
- Interpretation: most extracted value does not fit direct one-line runtime INI mapping and belongs in non-runtime lanes.

## SmartStat Lane Model

### runtime_ini_lane
- Count: **15** (0.33%)
- Meaning: Direct one-to-one runtime mappings and simple aliases compatible with governed INI sections.
- Authority class: `runtime-authoritative`
- Belongs:
  - one-to-one category->measure mappings
  - alias redirects to canonical runtime keys
  - high-confidence deterministic keys already aligned to CATEGORY_TO_MEASURE model
- Does NOT belong:
  - multi-measure expressions
  - grammar/context disambiguation
  - operator-assist-only reference phrases
  - ambiguous candidates without canonical confidence
### composite_expression_lane
- Count: **633** (13.81%)
- Meaning: Derived/multi-measure expression handling for ratios, splits, and text+token composition.
- Authority class: `future-architecture (non-runtime-authoritative today)`
- Belongs:
  - ratio expressions (K/BB, BB/K class)
  - slash-line structures
  - text-token hybrid templates
  - multi-token assembled outputs
- Does NOT belong:
  - single canonical mapping lines
  - simple aliases
  - runtime INI direct one-to-one entries
### semantic_resolver_lane
- Count: **1890** (41.24%)
- Meaning: Grammar-aware semantic interpretation and contextual intent resolution.
- Authority class: `future-architecture (non-runtime-authoritative today)`
- Belongs:
  - filter/qualifier grammar entries
  - playbook/query-aware semantic intent
  - context-dependent phrase disambiguation
- Does NOT belong:
  - direct static key->measure lines
  - pure operator UI help text only
  - uncertain phrases lacking deterministic interpretation
### tooling_operator_assist_lane
- Count: **499** (10.89%)
- Meaning: Reference/search/operator assistance and advisory surfaces without runtime authority.
- Authority class: `tooling-only/advisory`
- Belongs:
  - operator phrase help
  - candidate browsing/exploration
  - documentation and advisory UX surfaces
- Does NOT belong:
  - runtime mapping authority
  - automatic resolver decisions
  - mutation or execution behavior
### unresolved_or_ambiguous_lane
- Count: **1546** (33.73%)
- Meaning: Candidates with insufficient confidence or conflicting interpretations.
- Authority class: `blocked/pending manual review`
- Belongs:
  - manual-review-needed terms
  - conflicting semantic interpretations
  - low-confidence or context-fragile entries
- Does NOT belong:
  - automatic promotion
  - runtime mapping insertion
  - alias creation without clarity

## Bucket-to-Lane Mapping

- `runtime_mapping_candidate` -> `runtime_ini_lane`: Direct one-to-one mapping aligns with governed INI runtime mapping model.
- `alias_candidate` -> `runtime_ini_lane`: Alternate phrasing belongs to alias indirection over canonical runtime keys.
- `composite_candidate` -> `composite_expression_lane`: Requires composition/assembly beyond one-line static INI mapping.
- `semantic_layer_candidate` -> `semantic_resolver_lane`: Requires grammar/context interpretation beyond static mapping tables.
- `tooling_only` -> `tooling_operator_assist_lane`: Useful for operator/search/documentation support, not runtime authority.
- `ambiguous` -> `unresolved_or_ambiguous_lane`: Unsafe for promotion due to unresolved interpretation ambiguity.

## High-Value Candidate Clusters

### runtime_ini_lane
- Count: 15
- Not-yet-promoted high-confidence runtime candidates: 0
- Not-yet-promoted alias opportunities: 0
- Representative examples: singles, game_winning_rbi, go_ahead_rbi, GO-AHEAD RBI
### composite_expression_lane
- Count: 633
- `custom_header_placeholder_pattern`: 14
- `other_composite_pattern`: 291
- `ratio_or_slash_pattern`: 59
- `text_token_hybrid_pattern`: 295
- `vs_split_pattern`: 63
### semantic_resolver_lane
- Count: 1890
- `filter_grammar_cluster`: 28
- `playbook_metadata_cluster`: 1862
- Representative examples: filter grammar tokens (`season`, `home`, `away`), playbook metadata entries tied to basePage/writeToPattern context
### tooling_operator_assist_lane
- Count: 499
- Representative examples: batter_chase_percentage, batter_contact_percentage, batting_average, hits, doubles
### unresolved_or_ambiguous_lane
- Count: 1546
- Representative examples: rbi_opponent, at_bats_bases_empty, at_bats_bases_loaded, batter_foul_balls

## Why Only 4 Were Promoted

- Total classified candidates: **4583**.
- Runtime+alias lane volume: **15** (0.33%).
- Non-runtime lanes dominate: composite (633), semantic resolver (1890), tooling (499), ambiguous (1546).
- PASS_09 signoff approved exactly 4 lines; PASS_12 validated those 4 lines in runtime mapping file with correct placement and no collisions.
- Distinction breakdown:
  - Not lack of value: most extracted data remains valuable in non-runtime lanes.
  - Mismatch with current INI model: most entries are not one-line canonical mapping shapes.
  - Already-covered concepts: no additional high-confidence runtime mapping/alias opportunities remained beyond the promoted set.
  - Composite-only value: ratios/splits/text-token hybrids require a composite lane.
  - Semantic-layer value: grammar/context/playbook/filter handling requires resolver lane work.
  - Tooling/reference value: many entries are useful for operator assist but non-authoritative for runtime mapping.
  - Ambiguity: unresolved semantics remain fail-closed pending manual review.
- Conclusion: low promotion count reflects lane fit and governance safety, not lack of extracted semantic value.

## SmartStat Architecture Implications

- The INI mapping model is one runtime-authoritative lane, not the full semantic platform.
- Composite and semantic-resolver lanes carry substantial MLB OnAir value but require separate bounded contracts before implementation.
- Tooling/operator-assist lane can immediately support discoverability and workflow guidance without runtime authority expansion.
- Ambiguous lane remains fail-closed and should flow through controlled manual-review governance, not auto-promotion.

## Recommended Next Pass

- `MLB_ONAIR_COMPOSITE_EXPRESSION_CLUSTER_DEFINITION_PASS_03`
- Scope: read-only definition of deterministic composite-expression clusters and contract candidates from PASS_01 composite lane only.
- Boundary: no runtime mutation, no INI mutation, no resolver implementation.
