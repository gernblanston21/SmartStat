import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RuleTracePanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_trace_view"];
}

export function RuleTracePanel({ view }: RuleTracePanelProps): JSX.Element {
  return (
    <PanelShell title="Rule Evaluation Trace" subtitle="Traceability">
      <dl className="kv-grid">
        <dt>Projection Contract</dt>
        <dd>{view.projection_contract}</dd>
        <dt>Projection Kind</dt>
        <dd>{view.projection_kind}</dd>
        <dt>Input Artifact</dt>
        <dd>{view.input_artifact}</dd>
        <dt>Artifact Path</dt>
        <dd>{view.input_identity.artifact_path}</dd>
        <dt>Input Fingerprint SHA256</dt>
        <dd>{view.input_identity.input_fingerprint_sha256}</dd>
      </dl>
    </PanelShell>
  );
}
