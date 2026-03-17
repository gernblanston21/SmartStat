import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface IssuesPanelProps {
  view: RuntimePreviewViewModel["view_model"]["issues_view"];
}

export function IssuesPanel({ view }: IssuesPanelProps): JSX.Element {
  return (
    <PanelShell title="Issues Summary" subtitle="Preview Diagnostics">
      <dl className="kv-grid">
        <dt>Status</dt>
        <dd>{view.status}</dd>
        <dt>Error Count</dt>
        <dd>{view.error_count}</dd>
        <dt>Warning Count</dt>
        <dd>{view.warning_count}</dd>
      </dl>
      <h3>Errors</h3>
      {view.errors.length ? (
        <pre>{JSON.stringify(view.errors, null, 2)}</pre>
      ) : (
        <p>None</p>
      )}
      <h3>Warnings</h3>
      {view.warnings.length ? (
        <pre>{JSON.stringify(view.warnings, null, 2)}</pre>
      ) : (
        <p>None</p>
      )}
    </PanelShell>
  );
}
