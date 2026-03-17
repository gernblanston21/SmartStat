import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface ProjectionMetadataPanelProps {
  view: RuntimePreviewViewModel["view_model"]["projection_metadata_view"];
}

export function ProjectionMetadataPanel({
  view,
}: ProjectionMetadataPanelProps): JSX.Element {
  return (
    <PanelShell title="Projection Metadata" subtitle="Contract View">
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
      </dl>
    </PanelShell>
  );
}
