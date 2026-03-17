import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface ProjectionMetadataPanelProps {
  view: RuntimePreviewViewModel["view_model"]["projection_metadata_view"];
}

function compactHash(value: string): string {
  if (value.length <= 24) {
    return value;
  }
  return `${value.slice(0, 12)}...${value.slice(-12)}`;
}

export function ProjectionMetadataPanel({
  view,
}: ProjectionMetadataPanelProps): JSX.Element {
  const badges = [view.status_summary.status, "Contract View", "Traceability"];
  return (
    <PanelShell title="Projection Metadata" subtitle="Contract View" badges={badges}>
      <dl className="kv-grid">
        <dt>Projection Contract</dt>
        <dd>{view.projection_contract}</dd>
        <dt>Projection Kind</dt>
        <dd>{view.projection_kind}</dd>
        <dt>Input Artifact</dt>
        <dd>{view.input_artifact}</dd>
        <dt>Artifact Path</dt>
        <dd>{view.input_identity.artifact_path}</dd>
        <dt>Input Fingerprint SHA256</dt>
        <dd>{view.input_identity.input_fingerprint_sha256}</dd>
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
