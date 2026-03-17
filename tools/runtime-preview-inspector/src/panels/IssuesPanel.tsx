import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface IssuesPanelProps {
  view: RuntimePreviewViewModel["view_model"]["issues_view"];
}

export function IssuesPanel({ view }: IssuesPanelProps): JSX.Element {
  const hasErrors = view.error_count > 0;
  const hasWarnings = view.warning_count > 0;
  const badges = [
    view.status,
    hasErrors ? "Errors Present" : "No Errors",
    hasWarnings ? "Warnings Present" : "No Warnings",
    "Preview",
    "Read-Only",
  ];
  return (
    <PanelShell
      title="Issues Summary"
      subtitle="Errors and warnings that affect preview confidence"
      badges={badges}
    >
      <p className="panel-note">
        If errors are present, treat the preview as needing review before trusting
        downstream interpretation.
      </p>
      <table className="summary-table">
        <thead>
          <tr>
            <th>Check</th>
            <th>Result</th>
            <th>Detail</th>
          </tr>
        </thead>
        <tbody>
          <tr className={view.status === "PASS" ? "is-pass" : "is-mismatch"}>
            <td>Preview Status</td>
            <td>{view.status}</td>
            <td>{view.status === "PASS" ? "No refusal state" : "Refusal state present"}</td>
          </tr>
          <tr className={hasErrors ? "is-mismatch" : "is-pass"}>
            <td>Error Count</td>
            <td>{view.error_count}</td>
            <td>{hasErrors ? "High-priority review needed" : "No errors detected"}</td>
          </tr>
          <tr className={hasWarnings ? "is-mismatch" : "is-pass"}>
            <td>Warning Count</td>
            <td>{view.warning_count}</td>
            <td>{hasWarnings ? "Advisory review recommended" : "No warnings detected"}</td>
          </tr>
        </tbody>
      </table>
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
