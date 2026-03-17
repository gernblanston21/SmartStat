import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RawPayloadPanelProps {
  view: RuntimePreviewViewModel["view_model"]["raw_payload_debug_view"];
}

export function RawPayloadPanel({ view }: RawPayloadPanelProps): JSX.Element {
  const badges = ["Read-Only", "Debug"];
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
