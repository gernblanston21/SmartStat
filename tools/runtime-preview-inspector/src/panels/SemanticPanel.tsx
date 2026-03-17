import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface SemanticPanelProps {
  view: RuntimePreviewViewModel["view_model"]["semantic_view"];
}

export function SemanticPanel({ view }: SemanticPanelProps): JSX.Element {
  const badges = ["Read-Only", "Contract View"];
  return (
    <PanelShell
      title="Semantic Interpretation"
      subtitle="How the preview interpreted scope and evidence"
      badges={badges}
    >
      <p className="panel-note">
        This tells you what context the preview settled on (for example, a scope
        like career or season) and where that interpretation came from.
      </p>
      <dl className="kv-grid">
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
