# MLB Composite Advisory Test Vector Report (PASS_06)

## Scope
- Pass: `MLB_ONAIR_COMPOSITE_ADVISORY_TEST_VECTOR_PASS_06`
- Mode: read-only test-vector analysis
- In-scope families only:
  - `ratio_or_slash_cluster`
  - `multi_measure_sequence_cluster`
- This pass did **not** modify runtime code, INI files, or SmartStat execution surfaces.

## PASS_05 Spec Recap
PASS_06 validates the PASS_05 advisory model for:
- deterministic 8-step recognition pipeline
- singular `PASS` / `DEFER` / `REJECT` outcomes
- deterministic reason-code assignment
- deterministic advisory output shape fields
- fail-closed behavior when structure/binding safety is not met

## Test Vector Model
Required vector fields:
- `test_id`
- `raw_expression`
- `normalized_expression`
- `expected_cluster_type`
- `expected_classification`
- `expected_reason_code`
- `expected_confidence_level`
- `expected_extracted_components`
- `expected_canonical_candidates`
- `expected_interpretation_preview`
- `expected_notes_for_operator`
- `expected_pipeline_stop_step`
- `source_basis`

Organization model:
- cluster_family -> classification -> stable test_id ordering

Deterministic expectation rules:
- single expected classification per vector
- single expected reason code per vector
- ordered extracted components
- explicit pipeline stop step
- no alternate expected outputs

## Coverage Summary
- Total vectors: 27
- Precedence/conflict vectors: 4

By cluster:
- `ratio_or_slash_cluster`: PASS=3, DEFER=4, REJECT=4, total=11
- `multi_measure_sequence_cluster`: PASS=4, DEFER=4, REJECT=4, total=12

By classification (all vectors):
- PASS=7
- DEFER=9
- REJECT=11

By reason code:
- `DEFER_ALIAS_AMBIGUITY`: 1
- `DEFER_BOUNDARY_CUSTOM_PLACEHOLDER`: 2
- `DEFER_CONTEXT_DEPENDENT`: 1
- `DEFER_UNRESOLVED_COMPONENT`: 5
- `PASS_FULLY_RESOLVED`: 7
- `REJECT_EMPTY_EXPRESSION`: 1
- `REJECT_INVALID_COMPONENT_COUNT`: 2
- `REJECT_INVALID_DELIMITER`: 2
- `REJECT_MIXED_CLUSTER_SIGNAL`: 2
- `REJECT_PATH_LIKE_SYNTAX`: 3
- `REJECT_UNKNOWN_STRUCTURE`: 1

Edge-case coverage:
- delimiter edge cases
- component-count bounds
- path-like false-positive rejection
- mixed-cluster signal rejection
- unresolved component defer
- alias ambiguity defer
- custom placeholder boundary defer
- template-context defer
- ordering-sensitive sequences

## Representative Vectors
PASS examples:
- `RATIO_PASS_001` -> `K/BB` -> `PASS_FULLY_RESOLVED`
- `RATIO_PASS_003` -> `STRIKEOUTS/BB` -> alias-to-canonical PASS
- `SEQUENCE_PASS_001` -> `AVG/HR/RBI` -> ordered 3-component PASS

DEFER examples:
- `RATIO_DEFER_001` -> `K/UNKNOWN` -> `DEFER_UNRESOLVED_COMPONENT`
- `RATIO_DEFER_002` -> `AVG/K` (context_unspecified) -> `DEFER_ALIAS_AMBIGUITY`
- `SEQUENCE_DEFER_003` -> `AVG,HR,{{custom.header_rbi}}` -> `DEFER_BOUNDARY_CUSTOM_PLACEHOLDER`

REJECT examples:
- `RATIO_REJECT_001` -> `K//BB` -> `REJECT_INVALID_DELIMITER`
- `RATIO_REJECT_003` -> path-like expression -> `REJECT_PATH_LIKE_SYNTAX`
- `SEQUENCE_REJECT_003` -> `AVG;HR;RBI` -> `REJECT_UNKNOWN_STRUCTURE`

## Precedence / Conflict Coverage
The `PRECEDENCE_*` vectors explicitly protect against:
- invalid structure being misclassified as PASS
- unresolved component defer being bypassed to PASS
- mixed-signal inputs downgrading incorrectly to DEFER/PASS
- path-like content reaching delimiter parsing
- empty-input not terminating at intake

## Why This Matters
This corpus creates a deterministic conformance surface before implementation.
It ensures future advisory-only recognition logic can be validated for:
- correct cluster assignment
- correct reason-code precedence
- correct fail-closed behavior
- correct ordered extraction and advisory explanation shape

Without this corpus, future implementation could drift silently in classification behavior.

## Recommended Next Pass
- `MLB_ONAIR_COMPOSITE_ADVISORY_VECTOR_REVIEW_SIGNOFF_PASS_07`
- Bounded objective: perform conservative signoff on PASS_06 vectors and freeze a minimal approved conformance subset for future advisory-only implementation validation.
- No runtime mutation or INI mutation in that pass.
