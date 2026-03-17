import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { ComparisonTable } from "./ComparisonTable";
import { PanelShell } from "./PanelShell";
import { compactHash, compactPath } from "./valueFormatting";

interface ProjectionMetadataPanelProps {
  view: RuntimePreviewViewModel["view_model"]["projection_metadata_view"];
  traceView: RuntimePreviewViewModel["view_model"]["rule_evaluation_trace_view"];
}

export function ProjectionMetadataPanel({
  view,
  traceView,
}: ProjectionMetadataPanelProps): JSX.Element {
  const overlapRows = [
    {
      key: "projection_contract",
      field: "Projection Contract",
      left: view.projection_contract,
      right: traceView.projection_contract,
      pass: view.projection_contract === traceView.projection_contract,
    },
    {
      key: "projection_kind",
      field: "Projection Kind",
      left: view.projection_kind,
      right: traceView.projection_kind,
      pass: view.projection_kind === traceView.projection_kind,
    },
    {
      key: "input_artifact",
      field: "Input Artifact",
      left: compactPath(view.input_artifact),
      right: compactPath(traceView.input_artifact),
      leftTitle: view.input_artifact,
      rightTitle: traceView.input_artifact,
      pass: view.input_artifact === traceView.input_artifact,
    },
    {
      key: "artifact_path",
      field: "Artifact Path",
      left: compactPath(view.input_identity.artifact_path),
      right: compactPath(traceView.input_identity.artifact_path),
      leftTitle: view.input_identity.artifact_path,
      rightTitle: traceView.input_identity.artifact_path,
      pass: view.input_identity.artifact_path === traceView.input_identity.artifact_path,
    },
    {
      key: "input_fingerprint_sha256",
      field: "Input Fingerprint SHA256",
      left: compactHash(view.input_identity.input_fingerprint_sha256),
      right: compactHash(traceView.input_identity.input_fingerprint_sha256),
      leftTitle: view.input_identity.input_fingerprint_sha256,
      rightTitle: traceView.input_identity.input_fingerprint_sha256,
      pass:
        view.input_identity.input_fingerprint_sha256 ===
        traceView.input_identity.input_fingerprint_sha256,
    },
  ];
  const overlapPassCount = overlapRows.filter((row) => row.pass).length;
  const overlapMismatchCount = overlapRows.length - overlapPassCount;
  const badges = [
    view.status_summary.status,
    overlapMismatchCount === 0 ? "Overlap PASS" : "Overlap MISMATCH",
    "Contract View",
    "Traceability",
  ];
  return (
    <PanelShell
      title="Projection Metadata"
      subtitle="Identity, status, and source details for this preview"
      badges={badges}
    >
      <p className="panel-note">
        Use this section to verify whether this preview references the correct
        source artifact and identity values.
      </p>
      <p className="panel-note">
        Overlap check: {overlapPassCount}/{overlapRows.length} fields match rule trace.
      </p>
      <ComparisonTable
        leftLabel="Projection Metadata"
        rightLabel="Rule Trace"
        rows={overlapRows}
      />
      <dl className="kv-grid">
        <dt>Projection Contract</dt>
        <dd>{view.projection_contract}</dd>
        <dt>Projection Kind</dt>
        <dd>{view.projection_kind}</dd>
        <dt>Input Artifact</dt>
        <dd title={view.input_artifact}>{compactPath(view.input_artifact)}</dd>
        <dt>Artifact Path</dt>
        <dd title={view.input_identity.artifact_path}>
          {compactPath(view.input_identity.artifact_path)}
        </dd>
        <dt>Input Fingerprint SHA256</dt>
        <dd title={view.input_identity.input_fingerprint_sha256}>
          {compactHash(view.input_identity.input_fingerprint_sha256)}
        </dd>
        <dt>Status</dt>
        <dd>{view.status_summary.status}</dd>
        <dt>Error Count</dt>
        <dd>{view.status_summary.error_count}</dd>
        <dt>Warning Count</dt>
        <dd>{view.status_summary.warning_count}</dd>
        <dt>Normalized Plan Hash</dt>
        <dd title={view.deterministic_identity_summary.normalized_plan_hash}>
          {compactHash(view.deterministic_identity_summary.normalized_plan_hash)}
        </dd>
        <dt>Replay Identity</dt>
        <dd title={view.deterministic_identity_summary.replay_identity}>
          {compactHash(view.deterministic_identity_summary.replay_identity)}
        </dd>
        <dt>Validator Run Identity</dt>
        <dd title={view.deterministic_identity_summary.validator_run_identity}>
          {compactHash(view.deterministic_identity_summary.validator_run_identity)}
        </dd>
      </dl>
    </PanelShell>
  );
}
