# Viewer Runtime Preview Panel Mapping

Status: `SMARTSTAT_VIEWER_LAYER_IMPLEMENTATION_PASS_01`  
Scope: Contract mapping only. This document defines data surfaces, not UI design.

## Purpose

Define deterministic mapping from the frozen runtime preview payload into
viewer-facing inspection surfaces.

## Mapping Principles

1. Read-only consumption only.
2. Preserve deterministic ordering from contract surfaces.
3. Expose only contract-approved fields.
4. No derived runtime/apply outcomes.

## Adapter Output Sections

The adapter emits a single normalized view-model with these sections in order:

1. `intake_header`
2. `projection_metadata_view`
3. `semantic_view`
4. `issues_view`
5. `resolution_view`
6. `rule_evaluation_summary_view`
7. `rule_evaluation_trace_view`
8. `deterministic_identity_view`
9. `raw_payload_debug_view`

## Panel-to-Source Mapping

### 1) Intake Header

Source:

1. envelope: `status`, `slice_name`, `preview_kind`, `provider_mode`,
   `normalization_rule`, `page_name`, `page_template`,
   `supported_read_surfaces`
2. payload: `bridge_mode`, `mutation_authorized`

### 2) Projection Metadata View

Source:

1. `preview_payload.projection_metadata`

### 3) Semantic Interpretation View

Source:

1. `preview_payload.projection_metadata.semantic_interpretation_summary`

### 4) Issues View

Source:

1. `preview_payload.projection_metadata.status_summary`
2. `preview_payload.projection_metadata.issues_summary`

### 5) Resolution View

Source:

1. `preview_payload.resolution_preview`

### 6) Rule Evaluation Summary View

Source:

1. `preview_payload.rule_evaluation_summary_preview.phase_order`
2. `preview_payload.rule_evaluation_summary_preview.ordered_rules`

### 7) Rule Evaluation Trace View

Source:

1. `preview_payload.rule_evaluation_trace_preview`

### 8) Deterministic Identity / Traceability View

Source:

1. `preview_payload.projection_metadata.deterministic_identity_summary`
2. `preview_payload.rule_evaluation_trace_preview.deterministic_identity_summary`

Constraint:

1. Both deterministic identity summaries must be byte-equivalent after
   normalization.

### 9) Raw Payload / Debug View

Source:

1. `payload_kind`
2. `tabfield_count`
3. `tabfield_order`
4. normalized `field_preview`
5. `projection_metadata`
6. `resolution_preview`
7. `rule_evaluation_summary_preview`
8. `rule_evaluation_trace_preview`

## Deterministic Ordering Rules

1. `projection_metadata` key order must remain:
   - `projection_contract`, `projection_kind`, `input_artifact`, `input_identity`,
     `status_summary`, `deterministic_identity_summary`,
     `semantic_interpretation_summary`, `issues_summary`
2. `rule_evaluation_summary_view.phase_order` must remain:
   - `STRUCTURAL`, `SEMANTIC`, `DETERMINISM`, `BOUNDARY`
3. `field_preview` display order must follow `tabfield_order`.
4. Adapter output object keys must remain stable in declared contract order.

## Non-Goals

This mapping does not include:

1. panel rendering implementation
2. runtime action controls
3. apply authorization states
4. Trio/socket integration controls
