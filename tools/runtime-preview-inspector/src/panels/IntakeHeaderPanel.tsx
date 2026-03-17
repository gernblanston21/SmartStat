import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface IntakeHeaderPanelProps {
  view: RuntimePreviewViewModel["view_model"]["intake_header"];
}

export function IntakeHeaderPanel({ view }: IntakeHeaderPanelProps): JSX.Element {
  return (
    <PanelShell title="Intake Header" subtitle="Read-Only Preview Contract View">
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
