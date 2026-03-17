# Viewer Runtime Preview Intake Contract

Status: `SMARTSTAT_VIEWER_LAYER_IMPLEMENTATION_PASS_01`  
Scope: Read-only tooling contract for runtime preview artifact intake only.

## Purpose

Define the strict intake contract for the Viewer Layer adapter that consumes the
frozen WP20 read-only preview artifact through Slice 02G.

This contract is non-authorizing and does not enable runtime/apply behavior.

## Authoritative Upstream Inputs

1. `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json`
2. `tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.json`
3. `docs/onair/plan-viewer-contract.md`
4. `docs/onair/wp20_runtime_lane_freeze_summary.md`

## Required Artifact Envelope Keys

The adapter intake requires these top-level keys:

1. `status`
2. `slice_name`
3. `preview_kind`
4. `provider_mode`
5. `normalization_rule`
6. `page_name`
7. `page_template`
8. `supported_read_surfaces`
9. `error_code`
10. `error_detail`
11. `preview_payload`

## Envelope Constraints

1. `status` must equal `success`.
2. `preview_kind` must equal `readonly_plan_bridge_preview_projection_intake_v1`.
3. `supported_read_surfaces` must be a string array.
4. `error_code` and `error_detail` must be present strings (empty strings allowed).

If any envelope constraint fails, intake must fail closed.

## Required `preview_payload` Keys

The adapter requires these `preview_payload` keys:

1. `payload_kind`
2. `bridge_mode`
3. `mutation_authorized`
4. `tabfield_count`
5. `tabfield_order`
6. `field_preview`
7. `projection_metadata`
8. `resolution_preview`
9. `rule_evaluation_summary_preview`
10. `rule_evaluation_trace_preview`

## Payload Constraints

1. `payload_kind` must equal `readonly_plan_bridge_preview_projection_intake_v1`.
2. `bridge_mode` must equal `read_only_preview`.
3. `mutation_authorized` must equal `false`.
4. `tabfield_count` must be a non-negative integer and must match:
   - `tabfield_order.length`
   - normalized `field_preview.length`
5. `tabfield_order` must be a deterministic string array.
6. `field_preview` must contain exactly one record for each tabfield name in
   `tabfield_order`.

## Required Projection Metadata Surface

Required key order for `projection_metadata`:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `status_summary`
6. `deterministic_identity_summary`
7. `semantic_interpretation_summary`
8. `issues_summary`

Required nested keys:

1. `input_identity.artifact_path`
2. `input_identity.input_fingerprint_sha256`
3. `status_summary.status`
4. `status_summary.error_count`
5. `status_summary.warning_count`
6. `deterministic_identity_summary.normalized_plan_hash`
7. `deterministic_identity_summary.replay_identity`
8. `deterministic_identity_summary.validator_run_identity`
9. `semantic_interpretation_summary.scope_resolution`
10. `semantic_interpretation_summary.effective_scope`
11. `semantic_interpretation_summary.evidence_source`
12. `issues_summary.errors`
13. `issues_summary.warnings`

## Required Resolution Surface

Required keys:

1. `status`
2. `scope_resolution`
3. `effective_scope`
4. `evidence_source`

`resolution_preview` values must match the corresponding
`projection_metadata.semantic_interpretation_summary` fields.

## Required Rule Evaluation Summary Surface

Required keys:

1. `phase_order`
2. `ordered_rules`

Constraints:

1. `phase_order` must equal:
   - `STRUCTURAL`, `SEMANTIC`, `DETERMINISM`, `BOUNDARY`
2. `ordered_rules` entries must each contain:
   - `category`
   - `rule_id`
   - `outcome`
3. Rule category ordering must be non-decreasing according to `phase_order`.

## Required Rule Evaluation Trace Surface

Required keys:

1. `projection_contract`
2. `projection_kind`
3. `input_artifact`
4. `input_identity`
5. `deterministic_identity_summary`

Overlap-equivalence constraints:

1. `projection_contract` must equal `projection_metadata.projection_contract`.
2. `projection_kind` must equal `projection_metadata.projection_kind`.
3. `input_artifact` must equal `projection_metadata.input_artifact`.
4. `input_identity` must equal `projection_metadata.input_identity`.
5. `deterministic_identity_summary` must equal
   `projection_metadata.deterministic_identity_summary`.

## Fail-Closed Rule

If any required key is missing, empty, malformed, mismatched, or non-deterministic:

1. Stop adaptation.
2. Raise contract error.
3. Emit no partial view-model.

## Explicit Non-Goals

This intake contract does not authorize:

1. runtime execution
2. apply behavior
3. Trio integration
4. socket integration
5. engine mutation
6. projection-contract expansion
