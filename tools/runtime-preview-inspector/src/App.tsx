import {
  EXPECTED_RULE_PHASE_ORDER,
  RulePhase,
  RuntimePreviewViewModel,
} from "./contracts/runtimePreviewIntake";
import { DeterministicIdentityPanel } from "./panels/DeterministicIdentityPanel";
import { IntakeHeaderPanel } from "./panels/IntakeHeaderPanel";
import { IssuesPanel } from "./panels/IssuesPanel";
import { PanelShell } from "./panels/PanelShell";
import { ProjectionMetadataPanel } from "./panels/ProjectionMetadataPanel";
import { RawPayloadPanel } from "./panels/RawPayloadPanel";
import { ResolutionPanel } from "./panels/ResolutionPanel";
import { RuleSummaryPanel } from "./panels/RuleSummaryPanel";
import { RuleTracePanel } from "./panels/RuleTracePanel";
import { SemanticPanel } from "./panels/SemanticPanel";

interface AppProps {
  viewModel: RuntimePreviewViewModel;
}

export const PANEL_RENDER_ORDER = [
  { key: "intake_header", title: "Intake Header" },
  { key: "projection_metadata_view", title: "Projection Metadata" },
  { key: "semantic_view", title: "Semantic Interpretation" },
  { key: "issues_view", title: "Issues Summary" },
  { key: "resolution_view", title: "Resolution Preview" },
  { key: "rule_evaluation_summary_view", title: "Rule Evaluation Summary" },
  { key: "rule_evaluation_trace_view", title: "Rule Evaluation Trace" },
  { key: "deterministic_identity_view", title: "Deterministic Identity" },
  { key: "raw_payload_debug_view", title: "Raw Payload Debug" },
] as const;

type PanelKey = (typeof PANEL_RENDER_ORDER)[number]["key"];
type ComparisonResult = "pass" | "mismatch" | "unavailable";
type PanelStatus = "ready" | "review" | "unavailable";
type ComparisonCheckKey =
  | "semantic_resolution"
  | "deterministic_identity"
  | "traceability_overlap"
  | "rule_phase_order";

const PANEL_TITLE_BY_KEY = Object.fromEntries(
  PANEL_RENDER_ORDER.map((panel) => [panel.key, panel.title])
) as Record<PanelKey, string>;

const UNAVAILABLE_LABEL = "Unavailable in this preview";
const NO_EVIDENCE_LABEL = "No evidence provided in current payload";

function comparisonResultLabel(result: ComparisonResult): "PASS" | "MISMATCH" | "UNAVAILABLE" {
  if (result === "pass") {
    return "PASS";
  }
  if (result === "mismatch") {
    return "MISMATCH";
  }
  return "UNAVAILABLE";
}

function comparisonResultClass(result: ComparisonResult): string {
  if (result === "pass") {
    return "is-pass";
  }
  if (result === "mismatch") {
    return "is-mismatch";
  }
  return "is-unavailable";
}

function panelStatusLabel(status: PanelStatus): "READY" | "REVIEW" | "UNAVAILABLE" {
  if (status === "ready") {
    return "READY";
  }
  if (status === "review") {
    return "REVIEW";
  }
  return "UNAVAILABLE";
}

function panelStatusClass(status: PanelStatus): string {
  if (status === "ready") {
    return "is-ready";
  }
  if (status === "review") {
    return "is-review";
  }
  return "is-unavailable";
}

function panelStatusNote(
  status: PanelStatus,
  reviewPairingKeys: readonly string[]
): string {
  if (status === "review" && reviewPairingKeys.length > 0) {
    return `Review evidence targeted by ${reviewPairingKeys.join(", ")}.`;
  }
  if (status === "unavailable") {
    return NO_EVIDENCE_LABEL;
  }
  return "Grounded evidence available.";
}

function comparisonPairingKey(key: ComparisonCheckKey): string {
  return `CHK-${key.replace(/_/g, "-").toUpperCase()}`;
}

function readText(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? value : undefined;
}

function readNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function readNonEmptyStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value) || value.length === 0) {
    return undefined;
  }
  const entries = value.filter((entry): entry is string => !!readText(entry));
  return entries.length === value.length ? entries : undefined;
}

function normalizeComparableValue(
  value: string | readonly string[] | undefined
): string | undefined {
  if (typeof value === "string") {
    return readText(value);
  }
  if (Array.isArray(value) && value.length > 0) {
    return `[${value.join(", ")}]`;
  }
  return undefined;
}

function compareValues(
  left: string | readonly string[] | undefined,
  right: string | readonly string[] | undefined
): ComparisonResult {
  const leftNormalized = normalizeComparableValue(left);
  const rightNormalized = normalizeComparableValue(right);

  if (!leftNormalized || !rightNormalized) {
    return "unavailable";
  }

  return leftNormalized === rightNormalized ? "pass" : "mismatch";
}

function combineComparisonResults(results: ComparisonResult[]): ComparisonResult {
  if (results.includes("mismatch")) {
    return "mismatch";
  }
  if (results.includes("unavailable")) {
    return "unavailable";
  }
  return "pass";
}

function isTruthyBoolean(value: unknown): value is boolean {
  return typeof value === "boolean";
}

interface InspectionSignal {
  key: string;
  label: string;
  result: ComparisonResult;
  detail: string;
  priority: "High" | "Medium";
}

interface ComparisonCheck {
  key: ComparisonCheckKey;
  label: string;
  result: ComparisonResult;
  priority: "High" | "Medium";
  pairingKey: string;
  leftTag: string;
  rightTag: string;
  evidenceTargets: PanelKey[];
}

interface DrillDownDiffLine {
  key: string;
  expectedLabel: string;
  expectedValue: string;
  actualLabel: string;
  actualValue: string;
  comparable: boolean;
  mismatch: boolean;
}

interface DrillDownCard {
  key: ComparisonCheckKey;
  label: string;
  result: ComparisonResult;
  priority: "High" | "Medium";
  pairingKey: string;
  leftTag: string;
  rightTag: string;
  evidenceTargets: PanelKey[];
  contributingFields: string[];
  diffLines: DrillDownDiffLine[];
  unavailableDiffCount: number;
}

function createDiffLine(
  key: string,
  expectedLabel: string,
  expectedValue: string | readonly string[] | undefined,
  actualLabel: string,
  actualValue: string | readonly string[] | undefined
): DrillDownDiffLine {
  const expected = normalizeComparableValue(expectedValue);
  const actual = normalizeComparableValue(actualValue);
  const comparable = !!expected && !!actual;

  return {
    key,
    expectedLabel,
    expectedValue: expected ?? UNAVAILABLE_LABEL,
    actualLabel,
    actualValue: actual ?? UNAVAILABLE_LABEL,
    comparable,
    mismatch: comparable ? expected !== actual : false,
  };
}

function hasIntakeHeaderEvidence(
  view: RuntimePreviewViewModel["view_model"]["intake_header"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.status) &&
    readText(view.slice_name) &&
    readText(view.preview_kind) &&
    readText(view.provider_mode) &&
    readText(view.normalization_rule) &&
    readText(view.page_name) &&
    readText(view.page_template) &&
    readText(view.bridge_mode) &&
    isTruthyBoolean(view.mutation_authorized) &&
    readNonEmptyStringArray(view.supported_read_surfaces)
  );
}

function hasProjectionMetadataEvidence(
  view: RuntimePreviewViewModel["view_model"]["projection_metadata_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.projection_contract) &&
    readText(view.projection_kind) &&
    readText(view.input_artifact) &&
    readText(view.input_identity.artifact_path) &&
    readText(view.input_identity.input_fingerprint_sha256) &&
    readText(view.status_summary.status) &&
    readNumber(view.status_summary.error_count) !== undefined &&
    readNumber(view.status_summary.warning_count) !== undefined &&
    readText(view.deterministic_identity_summary.normalized_plan_hash) &&
    readText(view.deterministic_identity_summary.replay_identity) &&
    readText(view.deterministic_identity_summary.validator_run_identity)
  );
}

function hasSemanticEvidence(
  view: RuntimePreviewViewModel["view_model"]["semantic_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.scope_resolution) &&
    readText(view.effective_scope) &&
    readText(view.evidence_source)
  );
}

function hasIssuesEvidence(
  view: RuntimePreviewViewModel["view_model"]["issues_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.status) &&
    readNumber(view.error_count) !== undefined &&
    readNumber(view.warning_count) !== undefined &&
    Array.isArray(view.errors) &&
    Array.isArray(view.warnings)
  );
}

function hasResolutionEvidence(
  view: RuntimePreviewViewModel["view_model"]["resolution_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.status) &&
    readText(view.scope_resolution) &&
    readText(view.effective_scope) &&
    readText(view.evidence_source)
  );
}

function hasRuleSummaryEvidence(
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_summary_view"] | undefined
): boolean {
  return !!(
    view &&
    readNonEmptyStringArray(view.phase_order) &&
    Array.isArray(view.ordered_rules) &&
    view.ordered_rules.length > 0 &&
    view.ordered_rules.every(
      (rule) => readText(rule.category) && readText(rule.rule_id) && readText(rule.outcome)
    )
  );
}

function hasRuleTraceEvidence(
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_trace_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.projection_contract) &&
    readText(view.projection_kind) &&
    readText(view.input_artifact) &&
    readText(view.input_identity.artifact_path) &&
    readText(view.input_identity.input_fingerprint_sha256) &&
    readText(view.deterministic_identity_summary.normalized_plan_hash) &&
    readText(view.deterministic_identity_summary.replay_identity) &&
    readText(view.deterministic_identity_summary.validator_run_identity)
  );
}

function hasDeterministicIdentityEvidence(
  view: RuntimePreviewViewModel["view_model"]["deterministic_identity_view"] | undefined
): boolean {
  return !!(
    view &&
    readText(view.projection_metadata.normalized_plan_hash) &&
    readText(view.projection_metadata.replay_identity) &&
    readText(view.projection_metadata.validator_run_identity) &&
    readText(view.rule_evaluation_trace.normalized_plan_hash) &&
    readText(view.rule_evaluation_trace.replay_identity) &&
    readText(view.rule_evaluation_trace.validator_run_identity)
  );
}

function hasRawPayloadEvidence(
  view: RuntimePreviewViewModel["view_model"]["raw_payload_debug_view"] | undefined
): boolean {
  const validFieldPreview =
    Array.isArray(view?.field_preview) &&
    view.field_preview.length > 0 &&
    view.field_preview.every(
      (entry) => readText(entry.name) && readText(entry.page_property) && readText(entry.custom_property)
    );

  return !!(
    view &&
    readText(view.payload_kind) &&
    readNumber(view.tabfield_count) !== undefined &&
    readNonEmptyStringArray(view.tabfield_order) &&
    validFieldPreview &&
    readText(view.resolution_preview.scope_resolution) &&
    Array.isArray(view.rule_evaluation_summary_preview.phase_order) &&
    Array.isArray(view.rule_evaluation_summary_preview.ordered_rules)
  );
}

function renderUnavailablePanel(title: string, subtitle: string): JSX.Element {
  return (
    <PanelShell
      title={title}
      subtitle={subtitle}
      badges={[UNAVAILABLE_LABEL, "Read-Only"]}
    >
      <p className="panel-note unavailable-panel-note">{UNAVAILABLE_LABEL}</p>
      <p className="panel-note unavailable-panel-note">{NO_EVIDENCE_LABEL}</p>
    </PanelShell>
  );
}

export default function App({ viewModel }: AppProps): JSX.Element {
  const sections =
    (viewModel?.view_model as Partial<RuntimePreviewViewModel["view_model"]>) ?? {};

  const intakeHeaderView = sections.intake_header;
  const projectionMetadataView = sections.projection_metadata_view;
  const semanticView = sections.semantic_view;
  const issuesView = sections.issues_view;
  const resolutionView = sections.resolution_view;
  const ruleSummaryView = sections.rule_evaluation_summary_view;
  const ruleTraceView = sections.rule_evaluation_trace_view;
  const deterministicIdentityView = sections.deterministic_identity_view;
  const rawPayloadView = sections.raw_payload_debug_view;

  const hasIntakeHeaderPanel = hasIntakeHeaderEvidence(intakeHeaderView);
  const hasProjectionMetadataPanel = hasProjectionMetadataEvidence(projectionMetadataView);
  const hasSemanticPanel = hasSemanticEvidence(semanticView) && hasResolutionEvidence(resolutionView);
  const hasIssuesPanel = hasIssuesEvidence(issuesView);
  const hasResolutionPanel = hasResolutionEvidence(resolutionView);
  const hasRuleSummaryPanel = hasRuleSummaryEvidence(ruleSummaryView);
  const hasRuleTracePanel = hasRuleTraceEvidence(ruleTraceView);
  const hasDeterministicIdentityPanel = hasDeterministicIdentityEvidence(deterministicIdentityView);
  const hasRawPayloadPanel = hasRawPayloadEvidence(rawPayloadView);

  const semanticScopeResolution = readText(semanticView?.scope_resolution);
  const semanticEffectiveScope = readText(semanticView?.effective_scope);
  const semanticEvidenceSource = readText(semanticView?.evidence_source);

  const resolutionScopeResolution = readText(resolutionView?.scope_resolution);
  const resolutionEffectiveScope = readText(resolutionView?.effective_scope);
  const resolutionEvidenceSource = readText(resolutionView?.evidence_source);

  const deterministicProjectionHash = readText(
    deterministicIdentityView?.projection_metadata.normalized_plan_hash
  );
  const deterministicProjectionReplay = readText(
    deterministicIdentityView?.projection_metadata.replay_identity
  );
  const deterministicProjectionRun = readText(
    deterministicIdentityView?.projection_metadata.validator_run_identity
  );

  const deterministicTraceHash = readText(
    deterministicIdentityView?.rule_evaluation_trace.normalized_plan_hash
  );
  const deterministicTraceReplay = readText(
    deterministicIdentityView?.rule_evaluation_trace.replay_identity
  );
  const deterministicTraceRun = readText(
    deterministicIdentityView?.rule_evaluation_trace.validator_run_identity
  );

  const projectionContract = readText(projectionMetadataView?.projection_contract);
  const projectionKind = readText(projectionMetadataView?.projection_kind);
  const projectionInputArtifact = readText(projectionMetadataView?.input_artifact);
  const projectionArtifactPath = readText(projectionMetadataView?.input_identity.artifact_path);
  const projectionFingerprint = readText(
    projectionMetadataView?.input_identity.input_fingerprint_sha256
  );

  const traceContract = readText(ruleTraceView?.projection_contract);
  const traceKind = readText(ruleTraceView?.projection_kind);
  const traceInputArtifact = readText(ruleTraceView?.input_artifact);
  const traceArtifactPath = readText(ruleTraceView?.input_identity.artifact_path);
  const traceFingerprint = readText(ruleTraceView?.input_identity.input_fingerprint_sha256);

  const phaseOrderActual = readNonEmptyStringArray(ruleSummaryView?.phase_order);
  const ruleCategorySequence =
    Array.isArray(ruleSummaryView?.ordered_rules) && ruleSummaryView.ordered_rules.length > 0
      ? Array.from(new Set(ruleSummaryView.ordered_rules.map((rule) => rule.category)))
      : undefined;

  const phaseOrderResult = compareValues(EXPECTED_RULE_PHASE_ORDER, phaseOrderActual);
  const semanticResolutionResult = combineComparisonResults([
    compareValues(semanticScopeResolution, resolutionScopeResolution),
    compareValues(semanticEffectiveScope, resolutionEffectiveScope),
    compareValues(semanticEvidenceSource, resolutionEvidenceSource),
  ]);
  const deterministicIdentityResult = combineComparisonResults([
    compareValues(deterministicProjectionHash, deterministicTraceHash),
    compareValues(deterministicProjectionReplay, deterministicTraceReplay),
    compareValues(deterministicProjectionRun, deterministicTraceRun),
  ]);
  const traceabilityOverlapResult = combineComparisonResults([
    compareValues(projectionContract, traceContract),
    compareValues(projectionKind, traceKind),
    compareValues(projectionInputArtifact, traceInputArtifact),
    compareValues(projectionArtifactPath, traceArtifactPath),
    compareValues(projectionFingerprint, traceFingerprint),
  ]);

  const previewStatus = readText(issuesView?.status);
  const previewStatusResult: ComparisonResult = !previewStatus
    ? "unavailable"
    : previewStatus === "PASS"
      ? "pass"
      : "mismatch";

  const errorCount = readNumber(issuesView?.error_count);
  const warningCount = readNumber(issuesView?.warning_count);

  const errorCountResult: ComparisonResult =
    errorCount === undefined ? "unavailable" : errorCount === 0 ? "pass" : "mismatch";
  const warningCountResult: ComparisonResult =
    warningCount === undefined ? "unavailable" : warningCount === 0 ? "pass" : "mismatch";

  const inspectionSignals: InspectionSignal[] = [
    {
      key: "preview_status",
      label: "Preview status",
      result: previewStatusResult,
      detail: previewStatus ?? UNAVAILABLE_LABEL,
      priority: "High",
    },
    {
      key: "error_count",
      label: "Error count",
      result: errorCountResult,
      detail: errorCount !== undefined ? String(errorCount) : UNAVAILABLE_LABEL,
      priority: "High",
    },
    {
      key: "semantic_resolution",
      label: "Semantic vs resolution",
      result: semanticResolutionResult,
      detail:
        semanticResolutionResult === "pass"
          ? "All key fields match"
          : semanticResolutionResult === "mismatch"
            ? "Scope/evidence mismatch"
            : UNAVAILABLE_LABEL,
      priority: "High",
    },
    {
      key: "deterministic_identity",
      label: "Deterministic identity parity",
      result: deterministicIdentityResult,
      detail:
        deterministicIdentityResult === "pass"
          ? "All identity fields match"
          : deterministicIdentityResult === "mismatch"
            ? "Identity mismatch"
            : UNAVAILABLE_LABEL,
      priority: "High",
    },
    {
      key: "traceability_overlap",
      label: "Projection vs trace overlap",
      result: traceabilityOverlapResult,
      detail:
        traceabilityOverlapResult === "pass"
          ? "Overlap fields aligned"
          : traceabilityOverlapResult === "mismatch"
            ? "Overlap mismatch"
            : UNAVAILABLE_LABEL,
      priority: "Medium",
    },
    {
      key: "rule_phase_order",
      label: "Rule phase order",
      result: phaseOrderResult,
      detail:
        phaseOrderResult === "pass"
          ? "Expected deterministic order"
          : phaseOrderResult === "mismatch"
            ? "Order differs from contract"
            : UNAVAILABLE_LABEL,
      priority: "Medium",
    },
    {
      key: "warning_count",
      label: "Warning count",
      result: warningCountResult,
      detail: warningCount !== undefined ? String(warningCount) : UNAVAILABLE_LABEL,
      priority: "Medium",
    },
  ];
  const reviewSignals = inspectionSignals.filter((signal) => signal.result !== "pass");
  const mismatchSignals = reviewSignals.filter((signal) => signal.result === "mismatch");
  const unavailableSignals = reviewSignals.filter((signal) => signal.result === "unavailable");
  const passSignals = inspectionSignals.length - reviewSignals.length;
  const highPriorityReviewSignals = inspectionSignals.filter(
    (signal) => signal.priority === "High" && signal.result !== "pass"
  );
  const mismatchCount = mismatchSignals.length;
  const unavailableCount = unavailableSignals.length;
  const reviewItemCount = reviewSignals.length;
  const highPriorityReviewCount = highPriorityReviewSignals.length;
  const overviewHealthy = reviewItemCount === 0;
  const reviewTone = overviewHealthy ? "Consistent Preview" : "Inconsistencies Found";
  const reviewToneClass = overviewHealthy
    ? "is-pass"
    : mismatchCount > 0
      ? "is-mismatch"
      : "is-unavailable";
  const topStatusToneLabel = overviewHealthy
    ? "All Core Checks PASS"
    : mismatchCount > 0
      ? "Review Needed"
      : "Evidence Unavailable";
  const comparisonChecks: ComparisonCheck[] = [
    {
      key: "semantic_resolution",
      label: "Semantic vs Resolution",
      result: semanticResolutionResult,
      priority: "High",
      pairingKey: comparisonPairingKey("semantic_resolution"),
      leftTag: "SEMANTIC",
      rightTag: "RESOLUTION",
      evidenceTargets: ["semantic_view", "resolution_view"],
    },
    {
      key: "deterministic_identity",
      label: "Deterministic Identity",
      result: deterministicIdentityResult,
      priority: "High",
      pairingKey: comparisonPairingKey("deterministic_identity"),
      leftTag: "PROJECTION",
      rightTag: "RULE TRACE",
      evidenceTargets: ["deterministic_identity_view"],
    },
    {
      key: "traceability_overlap",
      label: "Traceability Overlap",
      result: traceabilityOverlapResult,
      priority: "Medium",
      pairingKey: comparisonPairingKey("traceability_overlap"),
      leftTag: "PROJECTION",
      rightTag: "RULE TRACE",
      evidenceTargets: ["projection_metadata_view", "rule_evaluation_trace_view"],
    },
    {
      key: "rule_phase_order",
      label: "Rule Phase Order",
      result: phaseOrderResult,
      priority: "Medium",
      pairingKey: comparisonPairingKey("rule_phase_order"),
      leftTag: "EXPECTED",
      rightTag: "ACTUAL",
      evidenceTargets: ["rule_evaluation_summary_view"],
    },
  ];
  const hasPanelEvidenceByKey: Record<PanelKey, boolean> = {
    intake_header: hasIntakeHeaderPanel,
    projection_metadata_view: hasProjectionMetadataPanel && hasRuleTracePanel,
    semantic_view: hasSemanticPanel,
    issues_view: hasIssuesPanel,
    resolution_view: hasResolutionPanel,
    rule_evaluation_summary_view: hasRuleSummaryPanel,
    rule_evaluation_trace_view: hasRuleTracePanel,
    deterministic_identity_view: hasDeterministicIdentityPanel,
    raw_payload_debug_view: hasRawPayloadPanel,
  };

  const reviewPairingKeysByPanel = PANEL_RENDER_ORDER.reduce<Record<PanelKey, string[]>>(
    (acc, panel) => {
      acc[panel.key] = [];
      return acc;
    },
    {} as Record<PanelKey, string[]>
  );

  comparisonChecks
    .filter((check) => check.result === "mismatch")
    .forEach((check) => {
      check.evidenceTargets.forEach((panelKey) => {
        reviewPairingKeysByPanel[panelKey].push(check.pairingKey);
      });
    });

  const panelStatusByKey = PANEL_RENDER_ORDER.reduce<Record<PanelKey, PanelStatus>>(
    (acc, panel) => {
      if (!hasPanelEvidenceByKey[panel.key]) {
        acc[panel.key] = "unavailable";
        return acc;
      }
      if (reviewPairingKeysByPanel[panel.key].length > 0) {
        acc[panel.key] = "review";
        return acc;
      }
      acc[panel.key] = "ready";
      return acc;
    },
    {} as Record<PanelKey, PanelStatus>
  );

  const drillDownDetailsByKey: Record<
    ComparisonCheckKey,
    Pick<DrillDownCard, "contributingFields" | "diffLines" | "unavailableDiffCount">
  > = {
    semantic_resolution: {
      contributingFields: [
        "semantic.scope_resolution",
        "semantic.effective_scope",
        "semantic.evidence_source",
        "resolution.scope_resolution",
        "resolution.effective_scope",
        "resolution.evidence_source",
      ],
      diffLines: [
        createDiffLine(
          "scope_resolution",
          "semantic.scope_resolution",
          semanticScopeResolution,
          "resolution.scope_resolution",
          resolutionScopeResolution
        ),
        createDiffLine(
          "effective_scope",
          "semantic.effective_scope",
          semanticEffectiveScope,
          "resolution.effective_scope",
          resolutionEffectiveScope
        ),
        createDiffLine(
          "evidence_source",
          "semantic.evidence_source",
          semanticEvidenceSource,
          "resolution.evidence_source",
          resolutionEvidenceSource
        ),
      ],
      unavailableDiffCount: 0,
    },
    deterministic_identity: {
      contributingFields: [
        "projection_metadata.normalized_plan_hash",
        "projection_metadata.replay_identity",
        "projection_metadata.validator_run_identity",
        "rule_trace.normalized_plan_hash",
        "rule_trace.replay_identity",
        "rule_trace.validator_run_identity",
      ],
      diffLines: [
        createDiffLine(
          "normalized_plan_hash",
          "projection_metadata.normalized_plan_hash",
          deterministicProjectionHash,
          "rule_trace.normalized_plan_hash",
          deterministicTraceHash
        ),
        createDiffLine(
          "replay_identity",
          "projection_metadata.replay_identity",
          deterministicProjectionReplay,
          "rule_trace.replay_identity",
          deterministicTraceReplay
        ),
        createDiffLine(
          "validator_run_identity",
          "projection_metadata.validator_run_identity",
          deterministicProjectionRun,
          "rule_trace.validator_run_identity",
          deterministicTraceRun
        ),
      ],
      unavailableDiffCount: 0,
    },
    traceability_overlap: {
      contributingFields: [
        "projection_contract",
        "projection_kind",
        "input_artifact",
        "input_identity.artifact_path",
        "input_identity.input_fingerprint_sha256",
      ],
      diffLines: [
        createDiffLine(
          "projection_contract",
          "projection_metadata.projection_contract",
          projectionContract,
          "rule_trace.projection_contract",
          traceContract
        ),
        createDiffLine(
          "projection_kind",
          "projection_metadata.projection_kind",
          projectionKind,
          "rule_trace.projection_kind",
          traceKind
        ),
        createDiffLine(
          "input_artifact",
          "projection_metadata.input_artifact",
          projectionInputArtifact,
          "rule_trace.input_artifact",
          traceInputArtifact
        ),
        createDiffLine(
          "artifact_path",
          "projection_metadata.input_identity.artifact_path",
          projectionArtifactPath,
          "rule_trace.input_identity.artifact_path",
          traceArtifactPath
        ),
        createDiffLine(
          "input_fingerprint_sha256",
          "projection_metadata.input_identity.input_fingerprint_sha256",
          projectionFingerprint,
          "rule_trace.input_identity.input_fingerprint_sha256",
          traceFingerprint
        ),
      ],
      unavailableDiffCount: 0,
    },
    rule_phase_order: {
      contributingFields: [
        "rule_evaluation_summary.phase_order",
        "ordered_rules[*].category",
      ],
      diffLines: [
        createDiffLine(
          "phase_order",
          "expected.phase_order",
          EXPECTED_RULE_PHASE_ORDER,
          "actual.phase_order",
          phaseOrderActual
        ),
        createDiffLine(
          "rule_category_sequence",
          "expected.category_sequence",
          EXPECTED_RULE_PHASE_ORDER,
          "actual.category_sequence",
          ruleCategorySequence
        ),
      ],
      unavailableDiffCount: 0,
    },
  };
  const reviewDrillDownCards: DrillDownCard[] = comparisonChecks
    .filter((check) => check.result !== "pass")
    .map((check) => {
      const details = drillDownDetailsByKey[check.key];
      const unavailableDiffCount = details.diffLines.filter((line) => !line.comparable).length;
      const mismatchDiffLines = details.diffLines.filter((line) => line.mismatch);
      return {
        ...check,
        contributingFields: details.contributingFields,
        diffLines: mismatchDiffLines,
        unavailableDiffCount,
      };
    });

  const orderedRules =
    Array.isArray(ruleSummaryView?.ordered_rules) && ruleSummaryView.ordered_rules.length > 0
      ? ruleSummaryView.ordered_rules
      : undefined;

  const ruleCountsByPhase: Record<RulePhase, number | undefined> = {
    STRUCTURAL: undefined,
    SEMANTIC: undefined,
    DETERMINISM: undefined,
    BOUNDARY: undefined,
  };

  if (orderedRules) {
    const counts = orderedRules.reduce<Record<RulePhase, number>>(
      (acc, rule) => {
        acc[rule.category as RulePhase] += 1;
        return acc;
      },
      {
        STRUCTURAL: 0,
        SEMANTIC: 0,
        DETERMINISM: 0,
        BOUNDARY: 0,
      }
    );

    ruleCountsByPhase.STRUCTURAL = counts.STRUCTURAL;
    ruleCountsByPhase.SEMANTIC = counts.SEMANTIC;
    ruleCountsByPhase.DETERMINISM = counts.DETERMINISM;
    ruleCountsByPhase.BOUNDARY = counts.BOUNDARY;
  }

  const orderedPanels: Array<{
    key: PanelKey;
    className: string;
    status: PanelStatus;
    reviewPairingKeys: string[];
    node: JSX.Element;
  }> = [
    {
      key: "intake_header",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.intake_header,
      reviewPairingKeys: reviewPairingKeysByPanel.intake_header,
      node:
        hasIntakeHeaderPanel && intakeHeaderView
          ? <IntakeHeaderPanel view={intakeHeaderView} />
          : renderUnavailablePanel(
              "Intake Header",
              "Where this preview came from and what was read"
            ),
    },
    {
      key: "projection_metadata_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.projection_metadata_view,
      reviewPairingKeys: reviewPairingKeysByPanel.projection_metadata_view,
      node:
        hasProjectionMetadataPanel && hasRuleTracePanel && projectionMetadataView && ruleTraceView
          ? (
              <ProjectionMetadataPanel
                view={projectionMetadataView}
                traceView={ruleTraceView}
              />
            )
          : renderUnavailablePanel(
              "Projection Metadata",
              "Identity, status, and source details for this preview"
            ),
    },
    {
      key: "semantic_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.semantic_view,
      reviewPairingKeys: reviewPairingKeysByPanel.semantic_view,
      node:
        hasSemanticPanel && semanticView && resolutionView
          ? <SemanticPanel view={semanticView} resolutionView={resolutionView} />
          : renderUnavailablePanel(
              "Semantic Interpretation",
              "How the preview interpreted scope and evidence"
            ),
    },
    {
      key: "issues_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.issues_view,
      reviewPairingKeys: reviewPairingKeysByPanel.issues_view,
      node:
        hasIssuesPanel && issuesView
          ? <IssuesPanel view={issuesView} />
          : renderUnavailablePanel(
              "Issues Summary",
              "Errors and warnings that affect preview confidence"
            ),
    },
    {
      key: "resolution_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.resolution_view,
      reviewPairingKeys: reviewPairingKeysByPanel.resolution_view,
      node:
        hasResolutionPanel && resolutionView
          ? <ResolutionPanel view={resolutionView} />
          : renderUnavailablePanel(
              "Resolution Preview",
              "Final interpreted result shown in plain contract fields"
            ),
    },
    {
      key: "rule_evaluation_summary_view",
      className: "panel-slot panel-slot--full",
      status: panelStatusByKey.rule_evaluation_summary_view,
      reviewPairingKeys: reviewPairingKeysByPanel.rule_evaluation_summary_view,
      node:
        hasRuleSummaryPanel && ruleSummaryView
          ? <RuleSummaryPanel view={ruleSummaryView} />
          : renderUnavailablePanel(
              "Rule Evaluation Summary",
              "Order and outcomes of rule checks"
            ),
    },
    {
      key: "rule_evaluation_trace_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.rule_evaluation_trace_view,
      reviewPairingKeys: reviewPairingKeysByPanel.rule_evaluation_trace_view,
      node:
        hasRuleTracePanel && ruleTraceView
          ? <RuleTracePanel view={ruleTraceView} />
          : renderUnavailablePanel(
              "Rule Evaluation Trace",
              "Lineage fields used for traceability checks"
            ),
    },
    {
      key: "deterministic_identity_view",
      className: "panel-slot panel-slot--half",
      status: panelStatusByKey.deterministic_identity_view,
      reviewPairingKeys: reviewPairingKeysByPanel.deterministic_identity_view,
      node:
        hasDeterministicIdentityPanel && deterministicIdentityView
          ? <DeterministicIdentityPanel view={deterministicIdentityView} />
          : renderUnavailablePanel(
              "Deterministic Identity",
              "Side-by-side identity comparison"
            ),
    },
    {
      key: "raw_payload_debug_view",
      className: "panel-slot panel-slot--full",
      status: panelStatusByKey.raw_payload_debug_view,
      reviewPairingKeys: reviewPairingKeysByPanel.raw_payload_debug_view,
      node:
        hasRawPayloadPanel && rawPayloadView
          ? <RawPayloadPanel view={rawPayloadView} />
          : renderUnavailablePanel(
              "Raw Payload Debug",
              "Technical detail, with summary first"
            ),
    },
  ];

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>SmartStat Runtime Preview Inspector</h1>
        <p className="app-summary">
          This screen helps you review one saved preview file. It is read-only:
          nothing here writes to graphics systems, runs runtime actions, or changes
          SmartStat state.
        </p>
        <ul className="status-tags">
          <li>Read-Only</li>
          <li>Preview</li>
          <li>Contract View</li>
          <li>Traceability</li>
          <li className={`status-tag--tone ${reviewToneClass}`}>{topStatusToneLabel}</li>
        </ul>
        <section className="overview-strip" aria-label="Screen overview">
          <article className="overview-card">
            <h3>What this screen is</h3>
            <p>
              A viewer for one frozen preview payload that shows whether important
              sections agree with each other.
            </p>
          </article>
          <article className="overview-card">
            <h3>What read-only means</h3>
            <p>
              You can inspect values, but you cannot run, apply, refresh, or mutate
              anything from this page.
            </p>
          </article>
          <article className="overview-card">
            <h3>How to review quickly</h3>
            <p>
              Start with high-priority checks, then use secondary context and panel
              detail only when something looks off.
            </p>
          </article>
        </section>
        <section className={`review-hero ${reviewToneClass}`} aria-label="Review outcome summary">
          <div className="review-hero-main">
            <h2>{reviewTone}</h2>
            <p>
              PASS checks: <strong>{passSignals}</strong> / {inspectionSignals.length}. Review
              items: <strong>{reviewItemCount}</strong>.
            </p>
            <h3>Discrepancy Summary</h3>
            {reviewItemCount ? (
              <ul className="compact-list">
                {reviewSignals.map((signal) => (
                  <li key={`mismatch-${signal.key}`}>
                    <strong>{signal.label}</strong>: {signal.detail}
                  </li>
                ))}
              </ul>
            ) : (
              <p className="panel-note">
                No discrepancy signals detected in this preview.
              </p>
            )}
          </div>
          <div className="review-hero-metrics" aria-label="Review metrics">
            <article className={`metric-chip ${highPriorityReviewCount > 0 ? "is-mismatch" : "is-muted"}`}>
              <span>High-priority review items</span>
              <strong>{highPriorityReviewCount}</strong>
            </article>
            <article className={`metric-chip ${reviewItemCount > 0 ? "is-mismatch" : "is-muted"}`}>
              <span>Total review items</span>
              <strong>{reviewItemCount}</strong>
            </article>
            <article className={`metric-chip ${comparisonResultClass(errorCountResult)}`}>
              <span>Error count</span>
              <strong>{errorCount !== undefined ? errorCount : UNAVAILABLE_LABEL}</strong>
            </article>
            <article className={`metric-chip ${comparisonResultClass(warningCountResult)}`}>
              <span>Warning count</span>
              <strong>{warningCount !== undefined ? warningCount : UNAVAILABLE_LABEL}</strong>
            </article>
          </div>
        </section>
        <section className="check-first-strip" aria-label="What to check next">
          <h2>What to check next</h2>
          <p className={`check-first-result ${reviewToneClass}`}>{reviewTone}</p>
          <p className="check-first-note">
            Follow links in the last column to jump directly to detailed evidence panels.
          </p>
          <table className="summary-table summary-table--priority">
            <thead>
              <tr>
                <th>Comparison</th>
                <th>Result</th>
                <th>Where Detailed Evidence Lives</th>
              </tr>
            </thead>
            <tbody>
              {comparisonChecks.map((check) => (
                <tr
                  key={check.key}
                  className={`priority-row priority-${check.priority.toLowerCase()} ${
                    comparisonResultClass(check.result)
                  }`}
                >
                  <td>
                    <div className="comparison-label-cell">
                      <strong>{check.label}</strong>
                      <div className="comparison-cell-meta">
                        <span
                          className={`comparison-key-chip priority-${check.priority.toLowerCase()}`}
                        >
                          {check.pairingKey}
                        </span>
                      </div>
                    </div>
                  </td>
                  <td>
                    <span className={`parity-pill ${comparisonResultClass(check.result)}`}>
                      {comparisonResultLabel(check.result)}
                    </span>
                  </td>
                  <td>
                    <ul className="comparison-evidence-list">
                      {check.evidenceTargets.map((panelKey) => (
                        <li key={`${check.key}:${panelKey}`}>
                          <a href={`#panel-${panelKey}`}>{PANEL_TITLE_BY_KEY[panelKey]}</a>
                        </li>
                      ))}
                    </ul>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="check-first-hint">
            High-priority review items: <strong>{highPriorityReviewCount}</strong>. Mismatches:{" "}
            <strong>{mismatchCount}</strong>. Unavailable evidence: <strong>{unavailableCount}</strong>.
          </p>
          {reviewDrillDownCards.length > 0 ? (
            <section className="drilldown-strip" aria-label="Mismatch drill-down">
              <h3>Mismatch drill-down</h3>
              <p className="panel-note">
                Each card summarizes what failed, how values differ, and where detailed
                evidence lives.
              </p>
              <div className="drilldown-list">
                {reviewDrillDownCards.map((card) => (
                  <article
                    key={`drilldown-${card.key}`}
                    className={`drilldown-card ${
                      card.priority === "High" ? "is-high" : "is-medium"
                    } ${comparisonResultClass(card.result)}`}
                  >
                    <div className="drilldown-card-header">
                      <h4>
                        {card.label} - {card.result === "unavailable" ? "Evidence unavailable" : "Why this failed"}
                      </h4>
                      <div className="drilldown-card-meta">
                        <span
                          className={`comparison-key-chip priority-${card.priority.toLowerCase()}`}
                        >
                          {card.pairingKey}
                        </span>
                        <span className="drilldown-priority">{card.priority} Priority</span>
                        <span className={`drilldown-state-chip ${comparisonResultClass(card.result)}`}>
                          {comparisonResultLabel(card.result)}
                        </span>
                      </div>
                    </div>
                    {card.diffLines.length > 0 ? (
                      <>
                        <p className="drilldown-diff-count">
                          Differing grounded fields: <strong>{card.diffLines.length}</strong>
                        </p>
                        <ul className="drilldown-diff-list">
                          {card.diffLines.map((diffLine) => (
                            <li key={`diff-${card.key}-${diffLine.key}`} className="drilldown-diff-row">
                              <p className="drilldown-field-name">{diffLine.key}</p>
                              <div>
                                <span className="diff-label diff-label-before">{card.leftTag}:</span>{" "}
                                <code>{`${diffLine.expectedLabel} = ${diffLine.expectedValue}`}</code>
                              </div>
                              <div>
                                <span className="diff-label diff-label-after">{card.rightTag}:</span>{" "}
                                <code>{`${diffLine.actualLabel} = ${diffLine.actualValue}`}</code>
                              </div>
                            </li>
                          ))}
                        </ul>
                      </>
                    ) : (
                      <p className="panel-note unavailable-panel-note">{NO_EVIDENCE_LABEL}</p>
                    )}
                    {card.unavailableDiffCount > 0 ? (
                      <p className="panel-note unavailable-panel-note">
                        Some related fields are unavailable in this preview.
                      </p>
                    ) : null}
                    <p className="drilldown-fields-title">Key contributing fields</p>
                    <ul className="drilldown-fields">
                      {card.contributingFields.map((fieldName) => (
                        <li key={`field-${card.key}-${fieldName}`}>
                          <code>{fieldName}</code>
                        </li>
                      ))}
                    </ul>
                    <p className="drilldown-links-title">Related evidence panels:</p>
                    <ul className="comparison-evidence-list">
                      {card.evidenceTargets.map((panelKey) => (
                        <li key={`drilldown-link-${card.key}-${panelKey}`}>
                          <a href={`#panel-${panelKey}`}>{PANEL_TITLE_BY_KEY[panelKey]}</a>
                        </li>
                      ))}
                    </ul>
                  </article>
                ))}
              </div>
            </section>
          ) : null}
        </section>
        <details className="secondary-checks" open={!overviewHealthy}>
          <summary>Secondary checks and supporting context</summary>
          <p className="panel-note">
            Use this section after the high-priority checks. It provides additional
            parity and count context without crowding the first scan.
          </p>
          <p className="panel-note">
            If a top comparison fails, inspect related panels below in deterministic order.
            This context helps explain why a mismatch may exist.
          </p>
          <nav className="panel-jump-nav" aria-label="Panel order reference">
            <h3>Panel Order Reference</h3>
            <p className="panel-note">
              Status key: READY = grounded evidence available, REVIEW = linked to current mismatch
              checks, UNAVAILABLE = no evidence in this payload.
            </p>
            <ol className="panel-jump-list">
              {orderedPanels.map((panel) => (
                <li
                  key={`jump-${panel.key}`}
                  className={`panel-jump-item ${panelStatusClass(panel.status)}`}
                  data-panel-jump-key={panel.key}
                  data-panel-jump-status={panel.status}
                >
                  <a href={`#panel-${panel.key}`}>{PANEL_TITLE_BY_KEY[panel.key]}</a>
                  <span className={`panel-status-pill ${panelStatusClass(panel.status)}`}>
                    {panelStatusLabel(panel.status)}
                  </span>
                </li>
              ))}
            </ol>
          </nav>
          <section className="secondary-grid">
            <article className="overview-card">
              <h3>Rule Count Snapshot</h3>
              <table className="summary-table">
                <thead>
                  <tr>
                    <th>Phase</th>
                    <th>Rules</th>
                  </tr>
                </thead>
                <tbody>
                  {EXPECTED_RULE_PHASE_ORDER.map((phase) => (
                    <tr key={`phase-count-${phase}`}>
                      <td>{phase}</td>
                      <td>
                        {ruleCountsByPhase[phase] !== undefined
                          ? ruleCountsByPhase[phase]
                          : UNAVAILABLE_LABEL}
                      </td>
                    </tr>
                  ))}
                  <tr>
                    <td>Total Ordered Rules</td>
                    <td>{orderedRules ? orderedRules.length : UNAVAILABLE_LABEL}</td>
                  </tr>
                </tbody>
              </table>
            </article>
            <article className="overview-card">
              <h3>Traceability Snapshot</h3>
              <dl className="kv-grid compact-kv-grid">
                <dt>Projection Contract</dt>
                <dd>{projectionContract ?? UNAVAILABLE_LABEL}</dd>
                <dt>Projection Kind</dt>
                <dd>{projectionKind ?? UNAVAILABLE_LABEL}</dd>
                <dt>Input Artifact</dt>
                <dd>{projectionInputArtifact ?? UNAVAILABLE_LABEL}</dd>
              </dl>
            </article>
          </section>
        </details>
      </header>

      <main className="panel-stack" aria-label="Runtime Preview Panels">
        {orderedPanels.map((panel) => (
          <section
            key={panel.key}
            id={`panel-${panel.key}`}
            data-panel-key={panel.key}
            data-panel-status={panel.status}
            className={`${panel.className} panel-status-${panel.status}`}
          >
            <div className="panel-status-row">
              <span className={`panel-status-pill ${panelStatusClass(panel.status)}`}>
                {panelStatusLabel(panel.status)}
              </span>
              <span className="panel-status-note">
                {panelStatusNote(panel.status, panel.reviewPairingKeys)}
              </span>
            </div>
            {panel.node}
          </section>
        ))}
      </main>
    </div>
  );
}
