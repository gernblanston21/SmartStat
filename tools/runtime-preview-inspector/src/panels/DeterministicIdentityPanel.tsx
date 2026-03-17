import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface DeterministicIdentityPanelProps {
  view: RuntimePreviewViewModel["view_model"]["deterministic_identity_view"];
}

export function DeterministicIdentityPanel({
  view,
}: DeterministicIdentityPanelProps): JSX.Element {
  return (
    <PanelShell title="Deterministic Identity" subtitle="Traceability Contract View">
      <h3>Projection Metadata</h3>
      <dl className="kv-grid">
        <dt>Normalized Plan Hash</dt>
        <dd>{view.projection_metadata.normalized_plan_hash}</dd>
        <dt>Replay Identity</dt>
        <dd>{view.projection_metadata.replay_identity}</dd>
        <dt>Validator Run Identity</dt>
        <dd>{view.projection_metadata.validator_run_identity}</dd>
      </dl>

      <h3>Rule Evaluation Trace</h3>
      <dl className="kv-grid">
        <dt>Normalized Plan Hash</dt>
        <dd>{view.rule_evaluation_trace.normalized_plan_hash}</dd>
        <dt>Replay Identity</dt>
        <dd>{view.rule_evaluation_trace.replay_identity}</dd>
        <dt>Validator Run Identity</dt>
        <dd>{view.rule_evaluation_trace.validator_run_identity}</dd>
      </dl>
    </PanelShell>
  );
}
