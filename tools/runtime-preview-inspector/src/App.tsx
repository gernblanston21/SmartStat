import {
  EXPECTED_RULE_PHASE_ORDER,
  RuntimePreviewViewModel,
} from "./contracts/runtimePreviewIntake";
import { DeterministicIdentityPanel } from "./panels/DeterministicIdentityPanel";
import { IntakeHeaderPanel } from "./panels/IntakeHeaderPanel";
import { IssuesPanel } from "./panels/IssuesPanel";
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

function parityLabel(value: boolean): "PASS" | "MISMATCH" {
  return value ? "PASS" : "MISMATCH";
}

export default function App({ viewModel }: AppProps): JSX.Element {
  const sections = viewModel.view_model;
  const phaseOrderParity =
    JSON.stringify(sections.rule_evaluation_summary_view.phase_order) ===
    JSON.stringify(EXPECTED_RULE_PHASE_ORDER);
  const semanticResolutionParity =
    sections.semantic_view.scope_resolution === sections.resolution_view.scope_resolution &&
    sections.semantic_view.effective_scope === sections.resolution_view.effective_scope &&
    sections.semantic_view.evidence_source === sections.resolution_view.evidence_source;
  const deterministicIdentityParity =
    sections.deterministic_identity_view.projection_metadata.normalized_plan_hash ===
      sections.deterministic_identity_view.rule_evaluation_trace.normalized_plan_hash &&
    sections.deterministic_identity_view.projection_metadata.replay_identity ===
      sections.deterministic_identity_view.rule_evaluation_trace.replay_identity &&
    sections.deterministic_identity_view.projection_metadata.validator_run_identity ===
      sections.deterministic_identity_view.rule_evaluation_trace.validator_run_identity;
  const traceabilityOverlapParity =
    sections.projection_metadata_view.projection_contract ===
      sections.rule_evaluation_trace_view.projection_contract &&
    sections.projection_metadata_view.projection_kind ===
      sections.rule_evaluation_trace_view.projection_kind &&
    sections.projection_metadata_view.input_artifact ===
      sections.rule_evaluation_trace_view.input_artifact &&
    sections.projection_metadata_view.input_identity.artifact_path ===
      sections.rule_evaluation_trace_view.input_identity.artifact_path &&
    sections.projection_metadata_view.input_identity.input_fingerprint_sha256 ===
      sections.rule_evaluation_trace_view.input_identity.input_fingerprint_sha256;
  const topStatusPass = sections.issues_view.status === "PASS";
  const hasNoErrors = sections.issues_view.error_count === 0;
  const overviewHealthy =
    topStatusPass &&
    hasNoErrors &&
    semanticResolutionParity &&
    deterministicIdentityParity &&
    traceabilityOverlapParity &&
    phaseOrderParity;
  const reviewTone = overviewHealthy ? "Healthy Preview" : "Needs Review";
  const reviewToneClass = overviewHealthy ? "is-pass" : "is-mismatch";

  const orderedPanels: Array<{
    key: (typeof PANEL_RENDER_ORDER)[number]["key"];
    className: string;
    node: JSX.Element;
  }> =
    [
      {
        key: "intake_header",
        className: "panel-slot panel-slot--half",
        node: <IntakeHeaderPanel view={sections.intake_header} />,
      },
      {
        key: "projection_metadata_view",
        className: "panel-slot panel-slot--half",
        node: <ProjectionMetadataPanel view={sections.projection_metadata_view} />,
      },
      {
        key: "semantic_view",
        className: "panel-slot panel-slot--half",
        node: <SemanticPanel view={sections.semantic_view} />,
      },
      {
        key: "issues_view",
        className: "panel-slot panel-slot--half",
        node: <IssuesPanel view={sections.issues_view} />,
      },
      {
        key: "resolution_view",
        className: "panel-slot panel-slot--half",
        node: <ResolutionPanel view={sections.resolution_view} />,
      },
      {
        key: "rule_evaluation_summary_view",
        className: "panel-slot panel-slot--full",
        node: <RuleSummaryPanel view={sections.rule_evaluation_summary_view} />,
      },
      {
        key: "rule_evaluation_trace_view",
        className: "panel-slot panel-slot--half",
        node: <RuleTracePanel view={sections.rule_evaluation_trace_view} />,
      },
      {
        key: "deterministic_identity_view",
        className: "panel-slot panel-slot--half",
        node: <DeterministicIdentityPanel view={sections.deterministic_identity_view} />,
      },
      {
        key: "raw_payload_debug_view",
        className: "panel-slot panel-slot--full",
        node: <RawPayloadPanel view={sections.raw_payload_debug_view} />,
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
        </ul>
        <section className="overview-strip" aria-label="Screen overview">
          <article className="overview-card">
            <h3>What this screen is</h3>
            <p>
              A viewer for one frozen preview payload that shows whether key sections
              agree with each other.
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
            <h3>What you are seeing</h3>
            <p>
              Status, scope interpretation, rule ordering, traceability identities,
              and a compact raw snapshot.
            </p>
          </article>
        </section>
        <section className="check-first-strip" aria-label="What to check first">
          <h2>What to check first</h2>
          <p className={`check-first-result ${reviewToneClass}`}>{reviewTone}</p>
          <ol>
            <li>
              Preview status is <strong>{sections.issues_view.status}</strong>.
            </li>
            <li>
              Error count is <strong>{sections.issues_view.error_count}</strong> and
              warning count is <strong>{sections.issues_view.warning_count}</strong>.
            </li>
            <li>
              Core parity checks below should read <strong>PASS</strong> for a
              consistent preview.
            </li>
          </ol>
        </section>
        <section className="parity-strip" aria-label="Contract parity checks">
          <article className="parity-card">
            <h3>Semantic vs Resolution</h3>
            <p className={`parity-result ${semanticResolutionParity ? "is-pass" : "is-mismatch"}`}>
              {parityLabel(semanticResolutionParity)}
            </p>
          </article>
          <article className="parity-card">
            <h3>Deterministic Identity</h3>
            <p className={`parity-result ${deterministicIdentityParity ? "is-pass" : "is-mismatch"}`}>
              {parityLabel(deterministicIdentityParity)}
            </p>
          </article>
          <article className="parity-card">
            <h3>Traceability Overlap</h3>
            <p className={`parity-result ${traceabilityOverlapParity ? "is-pass" : "is-mismatch"}`}>
              {parityLabel(traceabilityOverlapParity)}
            </p>
          </article>
          <article className="parity-card">
            <h3>Rule Phase Order</h3>
            <p className={`parity-result ${phaseOrderParity ? "is-pass" : "is-mismatch"}`}>
              {parityLabel(phaseOrderParity)}
            </p>
          </article>
        </section>
      </header>

      <main className="panel-stack" aria-label="Runtime Preview Panels">
        {orderedPanels.map((panel) => (
          <section key={panel.key} data-panel-key={panel.key} className={panel.className}>
            {panel.node}
          </section>
        ))}
      </main>
    </div>
  );
}
