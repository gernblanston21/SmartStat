import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { EXPECTED_RULE_PHASE_ORDER } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";
import { compactRuleId } from "./valueFormatting";

interface RuleSummaryPanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_summary_view"];
}

export function RuleSummaryPanel({ view }: RuleSummaryPanelProps): JSX.Element {
  const phaseOrderMatches =
    JSON.stringify(view.phase_order) === JSON.stringify(EXPECTED_RULE_PHASE_ORDER);
  const categoryCounts = view.ordered_rules.reduce<Record<string, number>>((acc, rule) => {
    acc[rule.category] = (acc[rule.category] ?? 0) + 1;
    return acc;
  }, {});
  const badges = [phaseOrderMatches ? "Phase Order PASS" : "Phase Order MISMATCH", "Contract View"];
  return (
    <PanelShell
      title="Rule Evaluation Summary"
      subtitle="Order and outcomes of rule checks"
      badges={badges}
    >
      <p className="panel-note">
        Read this table top-to-bottom. It preserves deterministic ordering from
        the preview contract.
      </p>
      <h3>Phase Order</h3>
      <ol>
        {view.phase_order.map((phase) => (
          <li key={phase}>{phase}</li>
        ))}
      </ol>
      <h3>Category Distribution</h3>
      <ul className="compact-list">
        {view.phase_order.map((phase) => (
          <li key={`count-${phase}`}>
            <strong>{phase}</strong>: {categoryCounts[phase] ?? 0}
          </li>
        ))}
      </ul>

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
              <td title={rule.rule_id}>{compactRuleId(rule.rule_id)}</td>
              <td>{rule.outcome}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </PanelShell>
  );
}
