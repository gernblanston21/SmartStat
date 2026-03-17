import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { ComparisonTable } from "./ComparisonTable";
import { PanelShell } from "./PanelShell";
import { compactHash } from "./valueFormatting";

interface DeterministicIdentityPanelProps {
  view: RuntimePreviewViewModel["view_model"]["deterministic_identity_view"];
}

export function DeterministicIdentityPanel({
  view,
}: DeterministicIdentityPanelProps): JSX.Element {
  const parityRows = [
    {
      key: "normalized_plan_hash",
      field: "Normalized Plan Hash",
      left: compactHash(view.projection_metadata.normalized_plan_hash),
      right: compactHash(view.rule_evaluation_trace.normalized_plan_hash),
      leftTitle: view.projection_metadata.normalized_plan_hash,
      rightTitle: view.rule_evaluation_trace.normalized_plan_hash,
      pass:
        view.projection_metadata.normalized_plan_hash ===
        view.rule_evaluation_trace.normalized_plan_hash,
    },
    {
      key: "replay_identity",
      field: "Replay Identity",
      left: compactHash(view.projection_metadata.replay_identity),
      right: compactHash(view.rule_evaluation_trace.replay_identity),
      leftTitle: view.projection_metadata.replay_identity,
      rightTitle: view.rule_evaluation_trace.replay_identity,
      pass: view.projection_metadata.replay_identity === view.rule_evaluation_trace.replay_identity,
    },
    {
      key: "validator_run_identity",
      field: "Validator Run Identity",
      left: compactHash(view.projection_metadata.validator_run_identity),
      right: compactHash(view.rule_evaluation_trace.validator_run_identity),
      leftTitle: view.projection_metadata.validator_run_identity,
      rightTitle: view.rule_evaluation_trace.validator_run_identity,
      pass:
        view.projection_metadata.validator_run_identity ===
        view.rule_evaluation_trace.validator_run_identity,
    },
  ];
  const allParityPass = parityRows.every((row) => row.pass);
  const mismatchCount = parityRows.filter((row) => !row.pass).length;
  const badges = [allParityPass ? "Parity PASS" : "Parity MISMATCH", "Traceability"];

  return (
    <PanelShell
      title="Deterministic Identity"
      subtitle="Side-by-side identity comparison"
      badges={badges}
    >
      <p className="panel-note">
        All parity rows should read PASS. A mismatch means two contract surfaces
        disagree on identity.
      </p>
      <p className="panel-note">
        Mismatches: {mismatchCount} of {parityRows.length}.
      </p>
      <ComparisonTable
        leftLabel="Projection Metadata"
        rightLabel="Rule Trace"
        rows={parityRows}
      />
    </PanelShell>
  );
}
