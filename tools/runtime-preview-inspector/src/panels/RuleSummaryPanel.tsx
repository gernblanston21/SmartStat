import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RuleSummaryPanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_summary_view"];
}

export function RuleSummaryPanel({ view }: RuleSummaryPanelProps): JSX.Element {
  return (
    <PanelShell title="Rule Evaluation Summary" subtitle="Contract View">
      <h3>Phase Order</h3>
      <ol>
        {view.phase_order.map((phase) => (
          <li key={phase}>{phase}</li>
        ))}
      </ol>

      <h3>Ordered Rules</h3>
      <table>
        <thead>
          <tr>
            <th>Category</th>
            <th>Rule ID</th>
            <th>Outcome</th>
          </tr>
        </thead>
        <tbody>
          {view.ordered_rules.map((rule) => (
            <tr key={`${rule.category}:${rule.rule_id}`}>
              <td>{rule.category}</td>
              <td>{rule.rule_id}</td>
              <td>{rule.outcome}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </PanelShell>
  );
}
