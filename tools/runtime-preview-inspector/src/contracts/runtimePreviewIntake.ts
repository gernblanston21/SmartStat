export const RUNTIME_PREVIEW_INTAKE_CONTRACT =
  "viewer.runtime_preview_intake.v1" as const;
export const RUNTIME_PREVIEW_VIEW_MODEL_CONTRACT =
  "viewer.runtime_preview_view_model.v1" as const;

export const EXPECTED_PREVIEW_KIND =
  "readonly_plan_bridge_preview_projection_intake_v1" as const;
export const EXPECTED_PAYLOAD_KIND =
  "readonly_plan_bridge_preview_projection_intake_v1" as const;
export const EXPECTED_BRIDGE_MODE = "read_only_preview" as const;
export const EXPECTED_RULE_PHASE_ORDER = [
  "STRUCTURAL",
  "SEMANTIC",
  "DETERMINISM",
  "BOUNDARY",
] as const;

export type RuntimePreviewArtifactStatus = "success" | "fail_closed";
export type ValidationStatus = "PASS" | "REFUSE";
export type RulePhase = (typeof EXPECTED_RULE_PHASE_ORDER)[number];

export const REQUIRED_TRACEABILITY_OVERLAP_FIELDS = [
  "projection_contract",
  "projection_kind",
  "input_artifact",
  "input_identity",
  "deterministic_identity_summary",
] as const;

export interface InputIdentity {
  artifact_path: string;
  input_fingerprint_sha256: string;
}

export interface DeterministicIdentitySummary {
  normalized_plan_hash: string;
  replay_identity: string;
  validator_run_identity: string;
}

export interface StatusSummary {
  status: ValidationStatus;
  error_count: number;
  warning_count: number;
}

export interface SemanticInterpretationSummary {
  scope_resolution: string;
  effective_scope: string;
  evidence_source: string;
}

export interface IssuesSummary {
  errors: unknown[];
  warnings: unknown[];
}

export interface RuleSummaryEntry {
  category: RulePhase;
  rule_id: string;
  outcome: string;
}

export interface ProjectionMetadata {
  projection_contract: string;
  projection_kind: string;
  input_artifact: string;
  input_identity: InputIdentity;
  status_summary: StatusSummary;
  deterministic_identity_summary: DeterministicIdentitySummary;
  semantic_interpretation_summary: SemanticInterpretationSummary;
  issues_summary: IssuesSummary;
}

export interface ResolutionPreview {
  status: ValidationStatus;
  scope_resolution: string;
  effective_scope: string;
  evidence_source: string;
}

export interface RuleEvaluationSummaryPreview {
  phase_order: RulePhase[];
  ordered_rules: RuleSummaryEntry[];
}

export interface RuleEvaluationTracePreview {
  projection_contract: string;
  projection_kind: string;
  input_artifact: string;
  input_identity: InputIdentity;
  deterministic_identity_summary: DeterministicIdentitySummary;
}

export interface FieldPreviewEntry {
  name: string;
  page_property: string;
  custom_property: string;
}

export interface RuntimePreviewPayload {
  payload_kind: string;
  bridge_mode: string;
  mutation_authorized: boolean;
  tabfield_count: number;
  tabfield_order: string[];
  field_preview: FieldPreviewEntry[];
  projection_metadata: ProjectionMetadata;
  resolution_preview: ResolutionPreview;
  rule_evaluation_summary_preview: RuleEvaluationSummaryPreview;
  rule_evaluation_trace_preview: RuleEvaluationTracePreview;
}

export interface RuntimePreviewIntakeArtifact {
  status: RuntimePreviewArtifactStatus;
  slice_name: string;
  preview_kind: string;
  provider_mode: string;
  normalization_rule: string;
  page_name: string;
  page_template: string;
  supported_read_surfaces: string[];
  error_code: string;
  error_detail: string;
  preview_payload: RuntimePreviewPayload;
}

export interface RuntimePreviewIntakeHeaderView {
  status: RuntimePreviewArtifactStatus;
  slice_name: string;
  preview_kind: string;
  provider_mode: string;
  normalization_rule: string;
  page_name: string;
  page_template: string;
  supported_read_surfaces: string[];
  bridge_mode: string;
  mutation_authorized: boolean;
}

export interface RuntimePreviewIssuesView {
  status: ValidationStatus;
  error_count: number;
  warning_count: number;
  errors: unknown[];
  warnings: unknown[];
}

export interface RuntimePreviewDeterministicIdentityView {
  projection_metadata: DeterministicIdentitySummary;
  rule_evaluation_trace: DeterministicIdentitySummary;
}

export interface RuntimePreviewRawPayloadDebugView {
  payload_kind: string;
  tabfield_count: number;
  tabfield_order: string[];
  field_preview: FieldPreviewEntry[];
  projection_metadata: ProjectionMetadata;
  resolution_preview: ResolutionPreview;
  rule_evaluation_summary_preview: RuleEvaluationSummaryPreview;
  rule_evaluation_trace_preview: RuleEvaluationTracePreview;
}

export interface RuntimePreviewViewModelSections {
  intake_header: RuntimePreviewIntakeHeaderView;
  projection_metadata_view: ProjectionMetadata;
  semantic_view: SemanticInterpretationSummary;
  issues_view: RuntimePreviewIssuesView;
  resolution_view: ResolutionPreview;
  rule_evaluation_summary_view: RuleEvaluationSummaryPreview;
  rule_evaluation_trace_view: RuleEvaluationTracePreview;
  deterministic_identity_view: RuntimePreviewDeterministicIdentityView;
  raw_payload_debug_view: RuntimePreviewRawPayloadDebugView;
}

export interface RuntimePreviewViewModel {
  adapter_contract: typeof RUNTIME_PREVIEW_VIEW_MODEL_CONTRACT;
  source_preview_kind: string;
  source_payload_kind: string;
  view_model: RuntimePreviewViewModelSections;
}
