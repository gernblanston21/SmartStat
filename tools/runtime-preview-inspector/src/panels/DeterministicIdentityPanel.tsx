import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface DeterministicIdentityPanelProps {
  view: RuntimePreviewViewModel["view_model"]["deterministic_identity_view"];
}

function compactHash(value: string): string {
  if (value.length <= 24) {
    return value;
  }
  return `${value.slice(0, 12)}...${value.slice(-12)}`;
}

export function DeterministicIdentityPanel({
  view,
}: DeterministicIdentityPanelProps): JSX.Element {
  const parityRows = [
    {
      key: "normalized_plan_hash",
      left: view.projection_metadata.normalized_plan_hash,
      right: view.rule_evaluation_trace.normalized_plan_hash,
    },
    {
      key: "replay_identity",
      left: view.projection_metadata.replay_identity,
      right: view.rule_evaluation_trace.replay_identity,
    },
    {
      key: "validator_run_identity",
      left: view.projection_metadata.validator_run_identity,
      right: view.rule_evaluation_trace.validator_run_identity,
    },
  ];
  const allParityPass = parityRows.every((row) => row.left === row.right);
  const badges = [allParityPass ? "Parity PASS" : "Parity MISMATCH", "Traceability"];

  return (
    <PanelShell
      title="Deterministic Identity"
      subtitle="Traceability Contract View"
      badges={badges}
    >
      <h3>Projection Metadata</h3>
      <dl className="kv-grid">
        <dt>Normalized Plan Hash</dt>
        <dd title={view.projection_metadata.normalized_plan_hash}>
          {compactHash(view.projection_metadata.normalized_plan_hash)}
        </dd>
        <dt>Replay Identity</dt>
        <dd title={view.projection_metadata.replay_identity}>
          {compactHash(view.projection_metadata.replay_identity)}
        </dd>
        <dt>Validator Run Identity</dt>
        <dd title={view.projection_metadata.validator_run_identity}>
          {compactHash(view.projection_metadata.validator_run_identity)}
        </dd>
      </dl>

      <h3>Rule Evaluation Trace</h3>
      <dl className="kv-grid">
        <dt>Normalized Plan Hash</dt>
        <dd title={view.rule_evaluation_trace.normalized_plan_hash}>
          {compactHash(view.rule_evaluation_trace.normalized_plan_hash)}
        </dd>
        <dt>Replay Identity</dt>
        <dd title={view.rule_evaluation_trace.replay_identity}>
          {compactHash(view.rule_evaluation_trace.replay_identity)}
        </dd>
        <dt>Validator Run Identity</dt>
        <dd title={view.rule_evaluation_trace.validator_run_identity}>
          {compactHash(view.rule_evaluation_trace.validator_run_identity)}
        </dd>
      </dl>
      <h3>Field Parity</h3>
      <table>
        <thead>
          <tr>
            <th>Field</th>
            <th>Parity</th>
          </tr>
        </thead>
        <tbody>
          {parityRows.map((row) => (
            <tr key={row.key}>
              <td>{row.key}</td>
              <td>{row.left === row.right ? "PASS" : "MISMATCH"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </PanelShell>
  );
}
