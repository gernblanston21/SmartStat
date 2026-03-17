import { RuntimePreviewViewModel } from "./contracts/runtimePreviewIntake";
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

export default function App({ viewModel }: AppProps): JSX.Element {
  const sections = viewModel.view_model;

  const orderedPanels: Array<{ key: (typeof PANEL_RENDER_ORDER)[number]["key"]; node: JSX.Element }> =
    [
      {
        key: "intake_header",
        node: <IntakeHeaderPanel view={sections.intake_header} />,
      },
      {
        key: "projection_metadata_view",
        node: <ProjectionMetadataPanel view={sections.projection_metadata_view} />,
      },
      {
        key: "semantic_view",
        node: <SemanticPanel view={sections.semantic_view} />,
      },
      {
        key: "issues_view",
        node: <IssuesPanel view={sections.issues_view} />,
      },
      {
        key: "resolution_view",
        node: <ResolutionPanel view={sections.resolution_view} />,
      },
      {
        key: "rule_evaluation_summary_view",
        node: <RuleSummaryPanel view={sections.rule_evaluation_summary_view} />,
      },
      {
        key: "rule_evaluation_trace_view",
        node: <RuleTracePanel view={sections.rule_evaluation_trace_view} />,
      },
      {
        key: "deterministic_identity_view",
        node: <DeterministicIdentityPanel view={sections.deterministic_identity_view} />,
      },
      {
        key: "raw_payload_debug_view",
        node: <RawPayloadPanel view={sections.raw_payload_debug_view} />,
      },
    ];

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>SmartStat Runtime Preview Inspector</h1>
        <p>
          Read-Only viewer scaffold for frozen runtime preview payloads using
          adapter-mediated contract surfaces.
        </p>
        <ul className="status-tags">
          <li>Read-Only</li>
          <li>Preview</li>
          <li>Contract View</li>
          <li>Traceability</li>
        </ul>
      </header>

      <main className="panel-stack" aria-label="Runtime Preview Panels">
        {orderedPanels.map((panel) => (
          <section key={panel.key} data-panel-key={panel.key} className="panel-slot">
            {panel.node}
          </section>
        ))}
      </main>
    </div>
  );
}
