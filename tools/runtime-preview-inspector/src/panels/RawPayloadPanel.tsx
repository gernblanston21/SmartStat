import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import { PanelShell } from "./PanelShell";

interface RawPayloadPanelProps {
  view: RuntimePreviewViewModel["view_model"]["raw_payload_debug_view"];
}

export function RawPayloadPanel({ view }: RawPayloadPanelProps): JSX.Element {
  return (
    <PanelShell title="Raw Payload Debug" subtitle="Read-Only Debug Snapshot">
      <pre>{JSON.stringify(view, null, 2)}</pre>
    </PanelShell>
  );
}
