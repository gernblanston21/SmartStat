import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface IssuesPanelProps {
  view: RuntimePreviewViewModel["view_model"]["issues_view"];
}

export function IssuesPanel({ view }: IssuesPanelProps): JSX.Element {
  const badges = [view.status, "Preview", "Read-Only"];
  return (
    <PanelShell title="Issues Summary" subtitle="Preview Diagnostics" badges={badges}>
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
        <ul className="issue-list">
          {view.errors.map((error, index) => (
            <li key={`error-${index}`}>
              <code>{JSON.stringify(error)}</code>
            </li>
          ))}
        </ul>
      ) : (
        <p>None</p>
      )}
      <h3>Warnings</h3>
      {view.warnings.length ? (
        <ul className="issue-list">
          {view.warnings.map((warning, index) => (
            <li key={`warning-${index}`}>
              <code>{JSON.stringify(warning)}</code>
            </li>
          ))}
        </ul>
      ) : (
        <p>None</p>
      )}
    </PanelShell>
  );
}
