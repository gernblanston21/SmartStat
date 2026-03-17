import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RuleTracePanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_trace_view"];
}

function compactHash(value: string): string {
  if (value.length <= 24) {
    return value;
  }
  return `${value.slice(0, 12)}...${value.slice(-12)}`;
}

export function RuleTracePanel({ view }: RuleTracePanelProps): JSX.Element {
  const badges = ["Traceability", "Read-Only"];
  return (
    <PanelShell title="Rule Evaluation Trace" subtitle="Traceability" badges={badges}>
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
        <dd title={view.input_identity.input_fingerprint_sha256}>
          {compactHash(view.input_identity.input_fingerprint_sha256)}
        </dd>
      </dl>
      <h3>Overlap Surface Keys</h3>
      <ul className="compact-list">
        <li>projection_contract</li>
        <li>projection_kind</li>
        <li>input_artifact</li>
        <li>input_identity</li>
        <li>deterministic_identity_summary</li>
      </ul>
    </PanelShell>
  );
}
