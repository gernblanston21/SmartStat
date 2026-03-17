import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface ResolutionPanelProps {
  view: RuntimePreviewViewModel["view_model"]["resolution_view"];
}

export function ResolutionPanel({ view }: ResolutionPanelProps): JSX.Element {
  const badges = [view.status, "Preview", "Read-Only"];
  return (
    <PanelShell
      title="Resolution Preview"
      subtitle="Final interpreted result shown in plain contract fields"
      badges={badges}
    >
      <p className="panel-note">
        Compare this panel with Semantic Interpretation to confirm they describe
        the same outcome.
      </p>
      <table className="summary-table">
        <thead>
          <tr>
            <th>Resolution Check</th>
            <th>Value</th>
          </tr>
        </thead>
        <tbody>
          <tr className={view.status === "PASS" ? "is-pass" : "is-mismatch"}>
            <td>Status</td>
            <td>{view.status}</td>
          </tr>
          <tr>
            <td>Scope Resolution</td>
            <td>{view.scope_resolution}</td>
          </tr>
        </tbody>
      </table>
      <dl className="kv-grid">
        <dt>Status</dt>
        <dd>{view.status}</dd>
        <dt>Scope Resolution</dt>
        <dd>{view.scope_resolution}</dd>
        <dt>Effective Scope</dt>
        <dd>{view.effective_scope}</dd>
        <dt>Evidence Source</dt>
        <dd>{view.evidence_source}</dd>
      </dl>
    </PanelShell>
  );
}
