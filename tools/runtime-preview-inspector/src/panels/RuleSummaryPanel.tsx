import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { EXPECTED_RULE_PHASE_ORDER } from "../contracts/runtimePreviewIntake";
import { ComparisonTable } from "./ComparisonTable";
import { PanelShell } from "./PanelShell";
import { compactRuleId } from "./valueFormatting";

interface RuleSummaryPanelProps {
  view: RuntimePreviewViewModel["view_model"]["rule_evaluation_summary_view"];
}

export function RuleSummaryPanel({ view }: RuleSummaryPanelProps): JSX.Element {
  const phaseOrderMatches =
    JSON.stringify(view.phase_order) === JSON.stringify(EXPECTED_RULE_PHASE_ORDER);
  const phaseOrderRows = EXPECTED_RULE_PHASE_ORDER.map((expectedPhase, index) => {
    const actualPhase = view.phase_order[index] ?? "(missing)";
    return {
      key: `phase-${index}`,
      field: `Phase ${index + 1}`,
      left: expectedPhase,
      right: actualPhase,
      pass: expectedPhase === actualPhase,
    };
  });
  const categoryOutcomeCounts = EXPECTED_RULE_PHASE_ORDER.map((phase) => {
    const inPhase = view.ordered_rules.filter((rule) => rule.category === phase);
    const passCount = inPhase.filter((rule) => rule.outcome === "PASS").length;
    const nonPassCount = inPhase.length - passCount;
    return {
      phase,
      total: inPhase.length,
      pass: passCount,
      nonPass: nonPassCount,
    };
  });
  const badges = [
    phaseOrderMatches ? "Phase Order PASS" : "Phase Order MISMATCH",
    "Contract View",
    "Read-Only",
  ];
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
      <p className="panel-note">
        Total ordered rules: {view.ordered_rules.length}. Phase-order mismatches:{" "}
        {phaseOrderRows.filter((row) => !row.pass).length}.
      </p>
      <h3>Phase Order Comparison</h3>
      <ComparisonTable
        leftLabel="Expected"
        rightLabel="Preview"
        rows={phaseOrderRows}
      />

      <h3>Category and Outcome Counts</h3>
      <table>
        <thead>
          <tr>
            <th>Phase</th>
            <th>Rules</th>
            <th>PASS Outcomes</th>
            <th>Non-PASS Outcomes</th>
          </tr>
        </thead>
        <tbody>
          {categoryOutcomeCounts.map((entry) => (
            <tr key={`category-${entry.phase}`}>
              <td>{entry.phase}</td>
              <td>{entry.total}</td>
              <td>{entry.pass}</td>
              <td>{entry.nonPass}</td>
            </tr>
          ))}
          <tr>
            <td>Total</td>
            <td>{view.ordered_rules.length}</td>
            <td>
              {categoryOutcomeCounts.reduce((sum, entry) => sum + entry.pass, 0)}
            </td>
            <td>
              {categoryOutcomeCounts.reduce((sum, entry) => sum + entry.nonPass, 0)}
            </td>
          </tr>
        </tbody>
      </table>

      <h3>Ordered Rules</h3>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Category</th>
            <th>Rule ID</th>
            <th>Outcome</th>
          </tr>
        </thead>
        <tbody>
          {view.ordered_rules.map((rule, index) => (
            <tr key={`${rule.category}:${rule.rule_id}`}>
              <td>{index + 1}</td>
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
