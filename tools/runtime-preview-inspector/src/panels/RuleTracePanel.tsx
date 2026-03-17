import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";
import { compactHash, compactPath } from "./valueFormatting";

interface RuleTracePanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_trace_view"];
}

export function RuleTracePanel({ view }: RuleTracePanelProps): JSX.Element {
  const badges = ["Traceability", "Read-Only"];
  return (
    <PanelShell
      title="Rule Evaluation Trace"
      subtitle="Lineage fields used for traceability checks"
      badges={badges}
    >
      <p className="panel-note">
        These values are meant to match overlap fields in Projection Metadata.
      </p>
      <table className="summary-table">
        <thead>
          <tr>
            <th>Traceability Surface</th>
            <th>Value</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Overlap keys tracked</td>
            <td>5</td>
          </tr>
          <tr>
            <td>Deterministic identity fields</td>
            <td>3</td>
          </tr>
        </tbody>
      </table>
      <dl className="kv-grid">
        <dt>Projection Contract</dt>
        <dd>{view.projection_contract}</dd>
        <dt>Projection Kind</dt>
        <dd>{view.projection_kind}</dd>
        <dt>Input Artifact</dt>
        <dd title={view.input_artifact}>{compactPath(view.input_artifact)}</dd>
        <dt>Artifact Path</dt>
        <dd title={view.input_identity.artifact_path}>
          {compactPath(view.input_identity.artifact_path)}
        </dd>
        <dt>Input Fingerprint SHA256</dt>
        <dd title={view.input_identity.input_fingerprint_sha256}>
          {compactHash(view.input_identity.input_fingerprint_sha256)}
        </dd>
        <dt>Normalized Plan Hash</dt>
        <dd title={view.deterministic_identity_summary.normalized_plan_hash}>
          {compactHash(view.deterministic_identity_summary.normalized_plan_hash)}
        </dd>
        <dt>Replay Identity</dt>
        <dd title={view.deterministic_identity_summary.replay_identity}>
          {compactHash(view.deterministic_identity_summary.replay_identity)}
        </dd>
        <dt>Validator Run Identity</dt>
        <dd title={view.deterministic_identity_summary.validator_run_identity}>
          {compactHash(view.deterministic_identity_summary.validator_run_identity)}
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
