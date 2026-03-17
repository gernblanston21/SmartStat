import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { ComparisonTable } from "./ComparisonTable";
import { PanelShell } from "./PanelShell";

interface SemanticPanelProps {
  view: RuntimePreviewViewModel["view_model"]["semantic_view"];
  resolutionView: RuntimePreviewViewModel["view_model"]["resolution_view"];
}

export function SemanticPanel({
  view,
  resolutionView,
}: SemanticPanelProps): JSX.Element {
  const comparisonRows = [
    {
      key: "scope_resolution",
      field: "Scope Resolution",
      left: view.scope_resolution,
      right: resolutionView.scope_resolution,
      pass: view.scope_resolution === resolutionView.scope_resolution,
    },
    {
      key: "effective_scope",
      field: "Effective Scope",
      left: view.effective_scope,
      right: resolutionView.effective_scope,
      pass: view.effective_scope === resolutionView.effective_scope,
    },
    {
      key: "evidence_source",
      field: "Evidence Source",
      left: view.evidence_source,
      right: resolutionView.evidence_source,
      pass: view.evidence_source === resolutionView.evidence_source,
    },
  ];
  const mismatchCount = comparisonRows.filter((row) => !row.pass).length;
  const badges = [
    mismatchCount === 0 ? "Parity PASS" : "Parity MISMATCH",
    "Read-Only",
    "Contract View",
  ];
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
      <p className="panel-note">
        Comparison view: semantic values should match resolution values exactly.
      </p>
      <ComparisonTable
        leftLabel="Semantic"
        rightLabel="Resolution"
        rows={comparisonRows}
      />
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
