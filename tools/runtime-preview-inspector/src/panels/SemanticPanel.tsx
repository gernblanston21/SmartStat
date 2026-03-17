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
      subtitle="Read-Only Metadata"
      badges={badges}
    >
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
