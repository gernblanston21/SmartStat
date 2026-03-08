# OnAir Plan Capture Contract (WP-17)

Contract version: `wp17.plan_capture.v1`  
Status: `docs/tests/tooling only (read-only semantic architecture lane)`  
Last updated: `2026-03-08`

## 1. Purpose

This document defines the enforceable WP-17 captured-plan artifact contract for OnAir semantic architecture work.

A captured plan is a deterministic, read-only artifact that records slot-resolved semantic intent.

This contract does not execute runtime behavior.

## 2. Source Reconciliation

This contract is the enforceable reconciliation surface across:

- `docs/onair/slot-resolution-model.md`
- `docs/onair/plan-capture-shape.md`
- `docs/onair/plan-validation-model.md`

Interpretation policy:

- WP-17 capture records deterministic slot outcomes and refusal state.
- WP-18 validation enforces deeper semantic compatibility policies.
- WP-19 viewer renders captured/validated artifacts read-only.
- WP-20 runtime bridge is explicitly deferred.

## 3. What Captured Plan Is / Is Not

Captured plan is:

- a versioned JSON artifact
- deterministic and diffable
- read-only architecture/tooling output
- fail-closed when ambiguity/unsupported state exists

Captured plan is not:

- runtime apply behavior
- planner execution
- runtime bridge behavior
- resolver scoring/tie-policy logic

## 4. Artifact Types

The contract supports exactly two artifact types:

- `captured_plan`
- `capture_refusal`

Both types must include the same top-level keys:

1. `contract_version`
2. `artifact_type`
3. `mode`
4. `source`
5. `determinism`
6. `slot_sequence`
7. `terminal`
8. `refusal`
9. `deferred_boundaries`

No unknown top-level keys are allowed.

## 5. Deterministic Guarantees

Required deterministic guarantees:

- identical semantic inputs and same contract version produce equivalent canonical JSON artifacts
- slot order is explicit and contiguous (`1..N`)
- canonicalization version is explicit (`wp17.canonical_json.v1`)
- deterministic input fingerprint is required (`input_fingerprint_sha256`)

## 6. Allowed Field Shape (Normative)

Schema file:

- `docs/onair/plan-capture.schema.json`

Normative constants:

- `contract_version = wp17.plan_capture.v1`
- `mode = read_only_contract`
- `determinism.ordering_contract = slot_sequence.order_asc.v1`
- `determinism.canonicalization_version = wp17.canonical_json.v1`

`source` object (required keys only):

- `query_text`
- `normalized_query_text`
- `semantic_record_id`
- `semantic_record_type`
- `league` (`string` or `null`)

`slot_sequence` entry shape (required keys only):

- `order` (positive integer)
- `slot_class` (`family|operator|entity|scope|filter|terminal_measure|terminal_attribute|formatter`)
- `token`
- `token_kind` (`canonical|alias_normalized|literal_parameter`)
- `parameters` (`string[]`)
- `source_dictionary` (`query_skeletons|operator_grammar|entity_dictionary|filter_grammar_dictionary|measure_dictionary|attribute_dictionary|formatter_dictionary|derived`)
- `candidate_status` (`resolved|inferred`)

## 7. Capture Rules

For `captured_plan`:

- `refusal` must be `null`
- `terminal` must be non-null and match the terminal slot in `slot_sequence`
- first slot must be `family` or `operator`
- at least one `entity` slot must exist
- at most one `formatter` slot is allowed and it must be last
- no non-formatter slots may appear after terminal slot

For `capture_refusal`:

- `terminal` must be `null`
- `refusal` object is required with:
  - `stage = capture`
  - `code` in:
    - `AMBIGUOUS_SLOT`
    - `MISSING_TERMINAL`
    - `UNSUPPORTED_SLOT_CLASS`
    - `ORDER_VIOLATION`
    - `UNSUPPORTED_OPERATOR_FORM`
    - `UNSUPPORTED_STATE`
  - `message` (non-empty)
  - `blocking_slot_order` (`null` or valid slot order)

Fail-closed rule:

- ambiguous or unsupported capture state must be represented as `capture_refusal` and must not be emitted as `captured_plan`.

## 8. Deferred Boundaries (Required)

`deferred_boundaries` must include:

- `runtime apply behavior`
- `planner execution`
- `runtime bridge behavior`

## 9. Layer Boundaries

Capture layer (WP-17):

- records deterministic slot-sequence artifact shape only

Validation layer (WP-18, deferred):

- semantic compatibility policy enforcement
- richer terminal/entity/operator compatibility checks
- compatibility error taxonomy expansion

Viewer layer (WP-19, deferred):

- read-only visualization of capture + validation artifacts

Runtime bridge (WP-20, deferred):

- runtime integration and apply constraints

## 10. Non-Goals

This contract pass does not define or change:

- SmartStat runtime VBScript behavior
- resolver scoring/tie logic
- ambiguity policy behavior in runtime resolver
- transaction commit/write semantics
- production INI layouts

## 11. Change Control

Any behavior-affecting contract change must:

- bump contract version
- update schema
- update validator(s) under `tests/wp-17/`
- add deterministic fixture evidence and runner summary artifacts
