# MLB Composite Advisory Recognition Specification (PASS_05)

## Scope
- Pass: `MLB_ONAIR_COMPOSITE_ADVISORY_RECOGNITION_SPEC_PASS_05`
- Mode: read-only advisory-spec analysis
- In-scope clusters only:
  - `ratio_or_slash_cluster`
  - `multi_measure_sequence_cluster`
- Out of scope:
  - `text_plus_token_cluster`
  - `split_or_comparison_cluster`
  - `header_value_pair_cluster`
  - `contextual_template_cluster`
  - `unresolved_composite_cluster`
- Runtime/code/config boundary:
  - no runtime code modification
  - no INI mutation
  - no TrioCmd/apply/take/tabfield/socket behavior

## PASS_04 Recap
- Authoritative contract source: `tests/_scratch/onair-mlb-composite-contract-pass04/`
- PASS_04 in-scope membership counts:
  - `ratio_or_slash_cluster`: 161
    - pass: 130
    - defer: 30
    - reject: 1
  - `multi_measure_sequence_cluster`: 23
    - pass: 16
    - defer: 5
    - reject: 2
- PASS_05 objective:
  - define deterministic advisory recognition pipeline and output schema
  - keep all behavior advisory-only and non-authoritative

## Advisory Recognition Pipeline

### Step 1: Raw Expression Intake
- Input:
  - `raw_expression` string from an upstream candidate surface
  - `source_trace` metadata
- Output:
  - `intake_record` with immutable `raw_expression`
- Failure conditions:
  - empty/null expression -> `REJECT_EMPTY_EXPRESSION`

### Step 2: Normalization
- Input:
  - `intake_record.raw_expression`
- Deterministic normalization actions:
  - trim outer whitespace
  - collapse repeated internal spaces to single spaces
  - normalize case for recognition only (preserve raw for output)
  - normalize underscore-per form (`_per_`) to lexical `per` token for ratio checks
- Output:
  - `normalized_expression`
- Failure conditions:
  - normalization yields empty string -> `REJECT_EMPTY_EXPRESSION`

### Step 3: Cluster Identification
- Input:
  - `normalized_expression`
- Deterministic decision order:
  1. Detect path-like/file-target signatures first
  2. Detect ratio/slash shape
  3. Detect multi-measure sequence shape
  4. If multiple cluster signatures appear simultaneously, fail closed
- Output:
  - `cluster_type` in:
    - `ratio_or_slash_cluster`
    - `multi_measure_sequence_cluster`
    - `unsupported`
- Failure conditions:
  - unsupported structure -> `REJECT_UNKNOWN_STRUCTURE`
  - mixed cluster signal -> `REJECT_MIXED_CLUSTER_SIGNAL`
  - path-like syntax -> `REJECT_PATH_LIKE_SYNTAX`

### Step 4: Contract Validation (PASS_04 Rules)
- Input:
  - `cluster_type`
  - `normalized_expression`
- Validation rules:
  - `ratio_or_slash_cluster`
    - accepted separators: `/`, `per`
    - exactly 2 components only
    - no repeated components
    - no custom placeholder coupling
  - `multi_measure_sequence_cluster`
    - accepted shapes: comma-delimited sequence, or slash pair with parenthetical third metric
    - bounded components: 2..3
    - ordered sequence required
    - custom header placeholders push to defer
- Output:
  - `contract_validation_result`
- Failure conditions:
  - invalid delimiter -> `REJECT_INVALID_DELIMITER`
  - invalid component count -> `REJECT_INVALID_COMPONENT_COUNT`
  - unsupported mixed context -> `REJECT_UNSUPPORTED_MIXED_CONTEXT`

### Step 5: Component Extraction
- Input:
  - `normalized_expression`
  - `cluster_type`
- Output:
  - ordered `extracted_components[]`
  - `separator_family`
- Failure conditions:
  - component extraction ambiguity -> `DEFER_COMPONENT_PARSE_AMBIGUITY`

### Step 6: Canonical Binding Attempt (Read-Only)
- Input:
  - `extracted_components[]`
  - runtime reference mappings (`SmartStat_Mappings.ini`) read-only
- Binding rules:
  - attempt canonical match first
  - attempt alias-to-canonical fallback second
  - no mutation and no learning updates
- Output:
  - `canonical_candidates[]` aligned to extracted order
  - per-component binding status
- Failure conditions:
  - unresolved component(s) -> `DEFER_UNRESOLVED_COMPONENT`
  - ambiguous alias -> `DEFER_ALIAS_AMBIGUITY`

### Step 7: Classification Decision (PASS / DEFER / REJECT)
- Input:
  - cluster result
  - contract validation
  - extraction status
  - binding status
- Output:
  - `classification`: `PASS` or `DEFER` or `REJECT`
  - single deterministic `reason_code`
- Failure conditions:
  - none (this step must always terminate with one classification)

### Step 8: Advisory Output Construction
- Input:
  - all prior step outputs
- Output:
  - deterministic advisory object (see Advisory Output Structure)
- Failure conditions:
  - missing required advisory field -> `REJECT_ADVISORY_OUTPUT_INCOMPLETE`

## Classification Model

### PASS
Assign only when all are true:
- cluster recognized in-scope
- contract shape valid
- component extraction deterministic
- all components resolved to canonical candidates
- no ambiguity flags

Primary PASS reason code:
- `PASS_FULLY_RESOLVED`

Secondary PASS reason code:
- `PASS_STRUCTURAL_ONLY` (allowed only when contract/extraction are valid and canonical binding is intentionally unavailable in the current advisory environment, while still non-ambiguous)

### DEFER
Assign when structure is recognizable and safe to inspect, but not safe to finalize:
- unresolved component exists
- alias ambiguity exists
- boundary condition triggered (for example custom header coupling)
- template/context dependence enters scope

### REJECT
Assign when structural safety is not met:
- malformed delimiters
- unsupported shape
- mixed cluster signal
- path-like/file-target syntax
- empty or invalid normalized expression

## Reason Code System

| Reason Code | Classification | Generated At Step | Trigger |
|---|---|---|---|
| `PASS_FULLY_RESOLVED` | PASS | 7 | Cluster, contract, extraction, and all component bindings succeed with no ambiguity. |
| `PASS_STRUCTURAL_ONLY` | PASS | 7 | Structure is valid and deterministic but canonical binding is intentionally unavailable in advisory-only context. |
| `DEFER_UNRESOLVED_COMPONENT` | DEFER | 6 | One or more extracted components cannot be bound to canonical candidates. |
| `DEFER_ALIAS_AMBIGUITY` | DEFER | 6 | Alias binding yields more than one valid canonical target. |
| `DEFER_COMPONENT_PARSE_AMBIGUITY` | DEFER | 5 | Extraction produced multiple component parses under same cluster contract. |
| `DEFER_BOUNDARY_CUSTOM_PLACEHOLDER` | DEFER | 4 | Expression depends on custom placeholder/header coupling outside bounded cluster contract. |
| `DEFER_CONTEXT_DEPENDENT` | DEFER | 4 | Expression meaning depends on template context beyond PASS_04 contract scope. |
| `REJECT_EMPTY_EXPRESSION` | REJECT | 1 or 2 | Raw or normalized expression is empty. |
| `REJECT_INVALID_DELIMITER` | REJECT | 4 | Delimiter pattern violates cluster contract. |
| `REJECT_INVALID_COMPONENT_COUNT` | REJECT | 4 | Component count outside cluster bounds. |
| `REJECT_UNKNOWN_STRUCTURE` | REJECT | 3 | Expression does not match an in-scope composite cluster. |
| `REJECT_MIXED_CLUSTER_SIGNAL` | REJECT | 3 | Expression simultaneously matches conflicting cluster signatures. |
| `REJECT_PATH_LIKE_SYNTAX` | REJECT | 3 | Path/file-target syntax detected; not a stat composite expression. |
| `REJECT_UNSUPPORTED_MIXED_CONTEXT` | REJECT | 4 | Contract sees unsupported mixed context (cluster boundary violation). |
| `REJECT_ADVISORY_OUTPUT_INCOMPLETE` | REJECT | 8 | Required advisory output field could not be produced deterministically. |

Reason-code assignment precedence (deterministic):
1. `REJECT_EMPTY_EXPRESSION`
2. `REJECT_PATH_LIKE_SYNTAX`
3. `REJECT_MIXED_CLUSTER_SIGNAL`
4. `REJECT_UNKNOWN_STRUCTURE`
5. `REJECT_INVALID_DELIMITER`
6. `REJECT_INVALID_COMPONENT_COUNT`
7. `REJECT_UNSUPPORTED_MIXED_CONTEXT`
8. `DEFER_COMPONENT_PARSE_AMBIGUITY`
9. `DEFER_ALIAS_AMBIGUITY`
10. `DEFER_UNRESOLVED_COMPONENT`
11. `DEFER_BOUNDARY_CUSTOM_PLACEHOLDER`
12. `DEFER_CONTEXT_DEPENDENT`
13. `PASS_FULLY_RESOLVED`
14. `PASS_STRUCTURAL_ONLY`
15. `REJECT_ADVISORY_OUTPUT_INCOMPLETE`

## Advisory Output Structure

Exact field schema:
- `original_expression` (string)
- `normalized_expression` (string)
- `cluster_type` (enum):
  - `ratio_or_slash_cluster`
  - `multi_measure_sequence_cluster`
  - `unsupported`
- `classification` (enum): `PASS` | `DEFER` | `REJECT`
- `reason_code` (string)
- `extracted_components` (ordered string array)
- `canonical_candidates` (ordered array of objects):
  - `component` (string)
  - `canonical_key` (string or empty)
  - `binding_status` (enum): `resolved` | `unresolved` | `ambiguous`
- `interpretation_preview` (string)
- `confidence_level` (enum): `HIGH` | `MEDIUM` | `LOW`
- `notes_for_operator` (string)

Deterministic confidence mapping:
- `HIGH`:
  - classification `PASS`
  - reason `PASS_FULLY_RESOLVED`
- `MEDIUM`:
  - classification `DEFER`
  - reasons beginning with `DEFER_`
- `LOW`:
  - classification `REJECT`
  - reasons beginning with `REJECT_`

## Interpretation Rules

Allowed interpretation preview content:
- detected structure label
- ordered component list
- canonical candidate bindings (if available)
- bounded semantic phrasing only

Forbidden interpretation preview content:
- final numeric output values
- query execution or data retrieval results
- mutation-ready syntax
- apply/take/cue instructions
- runtime permissioning language

Cluster-specific interpretation templates:
- ratio/slash:
  - `Recognized ratio structure: <component_1> / <component_2>.`
- multi-measure sequence:
  - `Recognized ordered sequence: <component_1>, <component_2>[, <component_3>].`
- defer/reject must explicitly explain non-final status.

## Fail-Closed Rules
- Fail closed (DEFER/REJECT) when any of the following is true:
  - ambiguous delimiter usage
  - multiple valid component parses
  - unresolved components after canonical/alias lookup
  - mixed cluster signals
  - template-context dependency beyond in-scope contracts
  - custom header/value coupling for sequence recognition
- No fallback heuristic is allowed to auto-promote DEFER/REJECT to PASS.
- Same input expression must always produce the same classification and reason code.

## Operator Explanation Model

Deterministic operator message templates:
- PASS:
  - `Recognized as <cluster_label>. Components resolved: <component_summary>.`
- DEFER:
  - `Structure recognized as <cluster_label>, but advisory resolution is incomplete: <defer_reason>.`
- REJECT:
  - `Expression does not match supported composite contracts: <reject_reason>.`

Cluster labels:
- `ratio_or_slash_cluster` -> `ratio/slash composite`
- `multi_measure_sequence_cluster` -> `multi-measure sequence composite`

Reason text map (fixed):
- `PASS_FULLY_RESOLVED` -> `all components resolved`
- `PASS_STRUCTURAL_ONLY` -> `structure valid, canonical binding unavailable`
- `DEFER_UNRESOLVED_COMPONENT` -> `one or more components are unresolved`
- `DEFER_ALIAS_AMBIGUITY` -> `alias resolves to multiple canonical targets`
- `DEFER_COMPONENT_PARSE_AMBIGUITY` -> `multiple component parses detected`
- `DEFER_BOUNDARY_CUSTOM_PLACEHOLDER` -> `custom placeholder coupling exceeds contract boundary`
- `DEFER_CONTEXT_DEPENDENT` -> `template context required for safe interpretation`
- `REJECT_EMPTY_EXPRESSION` -> `expression is empty`
- `REJECT_INVALID_DELIMITER` -> `delimiter pattern is invalid`
- `REJECT_INVALID_COMPONENT_COUNT` -> `component count is out of bounds`
- `REJECT_UNKNOWN_STRUCTURE` -> `structure is not supported`
- `REJECT_MIXED_CLUSTER_SIGNAL` -> `conflicting structure signatures detected`
- `REJECT_PATH_LIKE_SYNTAX` -> `path-like syntax is not a composite stat expression`
- `REJECT_UNSUPPORTED_MIXED_CONTEXT` -> `mixed context is outside contract scope`
- `REJECT_ADVISORY_OUTPUT_INCOMPLETE` -> `advisory output could not be built safely`

## Examples

### PASS Examples
1. Expression: `K/BB`
- Classification: `PASS`
- Reason code: `PASS_FULLY_RESOLVED`
- Simplified advisory output:
  - cluster: `ratio_or_slash_cluster`
  - extracted components: [`K`, `BB`]
  - canonical candidates: [`strikeouts`, `walks`]
  - interpretation preview: `Recognized ratio structure: K / BB.`

2. Expression: `BB/K`
- Classification: `PASS`
- Reason code: `PASS_FULLY_RESOLVED`
- Simplified advisory output:
  - cluster: `ratio_or_slash_cluster`
  - extracted components: [`BB`, `K`]
  - canonical candidates: [`walks`, `strikeouts`]
  - interpretation preview: `Recognized ratio structure: BB / K.`

3. Expression: `AVG/HR/RBI`
- Classification: `PASS`
- Reason code: `PASS_FULLY_RESOLVED`
- Simplified advisory output:
  - cluster: `multi_measure_sequence_cluster`
  - extracted components: [`AVG`, `HR`, `RBI`]
  - canonical candidates: [`batting_average`, `home_runs`, `rbis`]
  - interpretation preview: `Recognized ordered sequence: AVG, HR, RBI.`

### DEFER Examples
1. Expression: `K/UNKNOWN`
- Classification: `DEFER`
- Reason code: `DEFER_UNRESOLVED_COMPONENT`
- Simplified advisory output:
  - cluster: `ratio_or_slash_cluster`
  - extracted components: [`K`, `UNKNOWN`]
  - canonical candidates: [`strikeouts`, unresolved]
  - interpretation preview: `Structure recognized as ratio/slash composite, but advisory resolution is incomplete: one or more components are unresolved.`

2. Expression: `AVG/HR/???`
- Classification: `DEFER`
- Reason code: `DEFER_UNRESOLVED_COMPONENT`
- Simplified advisory output:
  - cluster: `multi_measure_sequence_cluster`
  - extracted components: [`AVG`, `HR`, `???`]
  - canonical candidates: [`batting_average`, `home_runs`, unresolved]
  - interpretation preview: `Structure recognized as multi-measure sequence composite, but advisory resolution is incomplete: one or more components are unresolved.`

### REJECT Examples
1. Expression: `K//BB`
- Classification: `REJECT`
- Reason code: `REJECT_INVALID_DELIMITER`
- Simplified advisory output:
  - cluster: `unsupported`
  - extracted components: []
  - interpretation preview: `Expression does not match supported composite contracts: delimiter pattern is invalid.`

2. Expression: `E:/EDRIVE/MLB/HEADSHOTS/{{info.team.alias}}/{{info.player.last_name}}.png`
- Classification: `REJECT`
- Reason code: `REJECT_PATH_LIKE_SYNTAX`
- Simplified advisory output:
  - cluster: `unsupported`
  - extracted components: []
  - interpretation preview: `Expression does not match supported composite contracts: path-like syntax is not a composite stat expression.`

3. Expression: `AVG/HR, RBI`
- Classification: `REJECT`
- Reason code: `REJECT_MIXED_CLUSTER_SIGNAL`
- Simplified advisory output:
  - cluster: `unsupported`
  - extracted components: []
  - interpretation preview: `Expression does not match supported composite contracts: conflicting structure signatures detected.`

## Recommended Next Pass
- Recommended bounded next pass:
  - `MLB_ONAIR_COMPOSITE_ADVISORY_TEST_VECTOR_PASS_06`
- Objective:
  - Define a deterministic, read-only test vector corpus that maps expressions to expected `cluster_type`, `classification`, and `reason_code` under this PASS_05 spec.
- Boundary:
  - no runtime mutation
  - no INI mutation
  - no recognition implementation
