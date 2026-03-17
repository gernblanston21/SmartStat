import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface IntakeHeaderPanelProps {
  view: RuntimePreviewViewModel["view_model"]["intake_header"];
}

export function IntakeHeaderPanel({ view }: IntakeHeaderPanelProps): JSX.Element {
  const readOnlyPass = !view.mutation_authorized && view.bridge_mode === "read_only_preview";
  const badges = [
    view.status,
    readOnlyPass ? "Read-Only PASS" : "Read-Only MISMATCH",
    view.mutation_authorized ? "Mutation On" : "Mutation Off",
  ];
  return (
    <PanelShell
      title="Intake Header"
      subtitle="Where this preview came from and what was read"
      badges={badges}
    >
      <p className="panel-note">
        Start here if you want to confirm you are inspecting the expected page and
        expected preview mode.
      </p>
      <table className="summary-table">
        <thead>
          <tr>
            <th>Intake Check</th>
            <th>Result</th>
          </tr>
        </thead>
        <tbody>
          <tr className={view.status === "success" ? "is-pass" : "is-mismatch"}>
            <td>Artifact Status</td>
            <td>{view.status}</td>
          </tr>
          <tr className={readOnlyPass ? "is-pass" : "is-mismatch"}>
            <td>Read-Only Boundary</td>
            <td>{readOnlyPass ? "PASS" : "MISMATCH"}</td>
          </tr>
        </tbody>
      </table>
      <dl className="kv-grid">
        <dt>Status</dt>
        <dd>{view.status}</dd>
        <dt>Slice</dt>
        <dd>{view.slice_name}</dd>
        <dt>Preview Kind</dt>
        <dd>{view.preview_kind}</dd>
        <dt>Provider Mode</dt>
        <dd>{view.provider_mode}</dd>
        <dt>Normalization Rule</dt>
        <dd>{view.normalization_rule}</dd>
        <dt>Page</dt>
        <dd>
          {view.page_name} ({view.page_template})
        </dd>
        <dt>Bridge Mode</dt>
        <dd>{view.bridge_mode}</dd>
        <dt>Mutation Authorized</dt>
        <dd>{String(view.mutation_authorized)}</dd>
        <dt>Read Surfaces Count</dt>
        <dd>{view.supported_read_surfaces.length}</dd>
      </dl>
      <h3>Supported Read Surfaces</h3>
      <ul>
        {view.supported_read_surfaces.map((surface) => (
          <li key={surface}>{surface}</li>
        ))}
      </ul>
    </PanelShell>
  );
}
