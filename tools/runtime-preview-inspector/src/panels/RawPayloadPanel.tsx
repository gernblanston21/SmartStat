import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RawPayloadPanelProps {
  view: RuntimePreviewViewModel["view_model"]["raw_payload_debug_view"];
}

export function RawPayloadPanel({ view }: RawPayloadPanelProps): JSX.Element {
  const badges = ["Read-Only", "Debug"];
  const orderedRulesCount = view.rule_evaluation_summary_preview.ordered_rules.length;
  const phaseCount = view.rule_evaluation_summary_preview.phase_order.length;
  const semanticScope =
    view.projection_metadata.semantic_interpretation_summary.scope_resolution;
  const resolutionScope = view.resolution_preview.scope_resolution;
  const semanticResolutionParity = semanticScope === resolutionScope;
  return (
    <PanelShell
      title="Raw Payload Debug"
      subtitle="Technical detail, with summary first"
      badges={badges}
    >
      <p className="panel-note">
        Use this section when you need deeper inspection. The key checks are
        surfaced above first.
      </p>
      <table className="summary-table">
        <thead>
          <tr>
            <th>Snapshot Check</th>
            <th>Value</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Tabfields</td>
            <td>{view.tabfield_count}</td>
          </tr>
          <tr>
            <td>Rule Phases</td>
            <td>{phaseCount}</td>
          </tr>
          <tr>
            <td>Ordered Rules</td>
            <td>{orderedRulesCount}</td>
          </tr>
          <tr className={semanticResolutionParity ? "is-pass" : "is-mismatch"}>
            <td>Semantic vs Resolution Scope</td>
            <td>{semanticResolutionParity ? "PASS" : "MISMATCH"}</td>
          </tr>
        </tbody>
      </table>
      <dl className="kv-grid">
        <dt>Payload Kind</dt>
        <dd>{view.payload_kind}</dd>
        <dt>Tabfield Count</dt>
        <dd>{view.tabfield_count}</dd>
        <dt>Tabfield Order</dt>
        <dd>{view.tabfield_order.join(", ")}</dd>
      </dl>

      <h3>Field Preview (Contract-Ordered)</h3>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Page Property</th>
            <th>Custom Property</th>
          </tr>
        </thead>
        <tbody>
          {view.field_preview.map((entry) => (
            <tr key={entry.name}>
              <td>{entry.name}</td>
              <td>{entry.page_property}</td>
              <td>{entry.custom_property}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <details>
        <summary>View full raw payload JSON</summary>
        <pre>{JSON.stringify(view, null, 2)}</pre>
      </details>
    </PanelShell>
  );
}
