import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface ResolutionPanelProps {
  view: RuntimePreviewViewModel["view_model"]["resolution_view"];
}

export function ResolutionPanel({ view }: ResolutionPanelProps): JSX.Element {
  return (
    <PanelShell title="Resolution Preview" subtitle="Read-Only Resolution View">
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
