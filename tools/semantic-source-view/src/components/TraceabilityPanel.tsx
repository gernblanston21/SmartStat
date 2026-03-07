import React, { useMemo, useState } from "react";
import { TraceIndex, TreeSelection } from "../types";

interface TraceabilityPanelProps {
  traceIndex: TraceIndex;
  allowedRecordIds: Set<string>;
  sourceTypeFilter: string;
  search: string;
  selectedNode: TreeSelection | null;
}

function intersectsAny(left: string[], right: Set<string>): boolean {
  for (const value of left) {
    if (right.has(value)) {
      return true;
    }
  }
  return false;
}

export function TraceabilityPanel({
  traceIndex,
  allowedRecordIds,
  sourceTypeFilter,
  search,
  selectedNode,
}: TraceabilityPanelProps): JSX.Element {
  const [selectedRecordId, setSelectedRecordId] = useState<string | null>(null);
  const [selectedSourcePath, setSelectedSourcePath] = useState<string | null>(null);
  const q = search.trim().toLowerCase();

  const recordIds = useMemo(() => {
    return Object.keys(traceIndex.by_record_id).filter((recordId) => {
      if (!allowedRecordIds.has(recordId)) {
        return false;
      }
      if (q && !recordId.toLowerCase().includes(q)) {
        return false;
      }
      return true;
    });
  }, [traceIndex.by_record_id, allowedRecordIds, q]);

  const sourcePaths = useMemo(() => {
    return Object.keys(traceIndex.by_source_path).filter((sourcePath) => {
      const linked = traceIndex.by_source_path[sourcePath] ?? [];
      if (!intersectsAny(linked, allowedRecordIds)) {
        return false;
      }
      if (sourceTypeFilter !== "all" && !sourcePath.startsWith(`${sourceTypeFilter}/`)) {
        return false;
      }
      if (selectedNode?.sourcePath && sourcePath !== selectedNode.sourcePath) {
        return false;
      }
      if (q && !sourcePath.toLowerCase().includes(q)) {
        return false;
      }
      return true;
    });
  }, [traceIndex.by_source_path, allowedRecordIds, sourceTypeFilter, selectedNode, q]);

  const selectedRecordTrace = selectedRecordId ? traceIndex.by_record_id[selectedRecordId] : null;
  const selectedSourceRecords = selectedSourcePath ? traceIndex.by_source_path[selectedSourcePath] : null;

  return (
    <section className="trace-layout">
      <div className="trace-column">
        <header className="pane-header">
          <h3>By Record ID</h3>
        </header>
        <ul className="item-list compact">
          {recordIds.map((recordId) => (
            <li key={recordId}>
              <button
                type="button"
                className={selectedRecordId === recordId ? "item-button selected" : "item-button"}
                onClick={() => setSelectedRecordId(recordId)}
              >
                <span className="item-subtle">{recordId}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="trace-column">
        <header className="pane-header">
          <h3>By Source Path</h3>
        </header>
        <ul className="item-list compact">
          {sourcePaths.map((sourcePath) => (
            <li key={sourcePath}>
              <button
                type="button"
                className={selectedSourcePath === sourcePath ? "item-button selected" : "item-button"}
                onClick={() => setSelectedSourcePath(sourcePath)}
              >
                <span className="item-subtle">{sourcePath}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="detail-pane">
        <header className="pane-header">
          <h3>Trace Detail</h3>
        </header>
        {selectedRecordTrace ? (
          <pre>{JSON.stringify({ record_id: selectedRecordId, trace: selectedRecordTrace }, null, 2)}</pre>
        ) : selectedSourceRecords ? (
          <pre>{JSON.stringify({ source_path: selectedSourcePath, record_ids: selectedSourceRecords }, null, 2)}</pre>
        ) : (
          <p className="muted">Select a record id or source path to inspect traceability links.</p>
        )}
      </div>
    </section>
  );
}
