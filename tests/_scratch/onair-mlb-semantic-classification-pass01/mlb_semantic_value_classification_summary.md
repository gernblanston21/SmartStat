# MLB OnAir Semantic Value Classification Summary (PASS_01)

## Scope

- Pass: `MLB_ONAIR_SEMANTIC_VALUE_CLASSIFICATION_PASS_01`
- Mode: read-only classification/analysis only.
- Inputs analyzed: PASS_02 extraction set (via fallback docs/onair full artifacts), PASS_04/05/07/08/09 crosswalk/staging/signoff artifacts, plus runtime mapping references.
- No runtime changes and no INI changes were made.
- Path substitution recorded: Primary directory tests/_scratch/onair-mlb-pass02-full-index/ not found; used docs/onair PASS_02 full extraction artifacts.

## Classification Distribution

- `runtime_mapping_candidate`: 12 (0.26%)
- `alias_candidate`: 3 (0.07%)
- `composite_candidate`: 633 (13.81%)
- `semantic_layer_candidate`: 1890 (41.24%)
- `tooling_only`: 499 (10.89%)
- `ambiguous`: 1546 (33.73%)

## Key Findings

- Total classified candidates: **4583**.
- Runtime-safe one-to-one mapping candidates are a minority: **12 (0.26%)**.
- The dataset is dominated by non-direct layers: semantic-layer + composite + tooling = **3022 (65.94%)**.
- Ambiguity bucket remains present (**1546**) where prior passes held/rejected uncertain semantics.
- This explains why only a small controlled subset was safe for runtime promotion while most extracted value belongs outside direct INI mapping.

## Runtime Mapping Opportunities

- No additional high-confidence runtime_mapping_candidates were found beyond the already-promoted set under current deterministic rules.

## Alias Expansion Opportunities

- No strong new alias opportunities survived deterministic filtering without collision/redundancy.

## Composite Pattern Findings

- `custom_header_placeholder_pattern`: 14
- `other_composite_pattern`: 291
- `ratio_or_slash_pattern`: 59
- `text_token_hybrid_pattern`: 295
- `vs_split_pattern`: 63

## Semantic Layer Opportunities

- Semantic-layer candidate clusters (deterministic counts):
  - `filter_grammar_cluster`: 28
  - `playbook_metadata_cluster`: 1862
- These entries should not be forced into direct `CATEGORY_TO_MEASURE` runtime mapping and are better handled by resolver/grammar-aware semantic layers.

## Tooling Opportunities

- Tooling-only candidates (examples):
  - `batter_chase_percentage` (measure_catalog_full)
  - `batter_contact_percentage` (measure_catalog_full)
  - `batter_swing_percentage` (measure_catalog_full)
  - `batter_swings` (measure_catalog_full)
  - `batter_whiff_percentage` (measure_catalog_full)
  - `batting_average` (measure_catalog_full)
  - `walks` (measure_catalog_full)
  - `hits` (measure_catalog_full)
  - `doubles` (measure_catalog_full)
  - `triples` (measure_catalog_full)
- Tooling-only data can support operator UX, docs, and semantic inspection aids without changing runtime mapping behavior.

## Recommended Next Pass

- `MLB_ONAIR_SEMANTIC_LAYER_LANE_DEFINITION_PASS_02`: a bounded, read-only lane-definition pass that converts PASS_01 semantic/composite clusters into explicit non-runtime target lanes (semantic resolver grammar lane vs tooling lane), without INI/runtime mutation.
