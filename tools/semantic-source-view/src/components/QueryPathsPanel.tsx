import React, { useMemo } from "react";
import { notesToList } from "../data/viewModel";
import { QueryPathRecord } from "../types";

interface QueryPathsPanelProps {
  queryPaths: QueryPathRecord[];
  selectedQueryPathId: string | null;
  onSelectQueryPath: (pathId: string) => void;
}

function Row({ label, value }: { label: string; value: React.ReactNode }): JSX.Element {
  return (
    <div className="detail-row">
      <dt>{label}</dt>
      <dd>{value}</dd>
    </div>
  );
}

export function QueryPathsPanel({
  queryPaths,
  selectedQueryPathId,
  onSelectQueryPath,
}: QueryPathsPanelProps): JSX.Element {
  const selectedPath = useMemo(() => {
    if (!queryPaths.length) {
      return null;
    }
    return queryPaths.find((path) => path.id === selectedQueryPathId) ?? queryPaths[0];
  }, [queryPaths, selectedQueryPathId]);

  return (
    <section className="split-panel">
      <div className="list-pane">
        <header className="pane-header">
          <h3>Query Paths ({queryPaths.length})</h3>
        </header>
        <ul className="item-list">
          {queryPaths.map((path) => (
            <li key={path.id}>
              <button
                type="button"
                className={selectedPath?.id === path.id ? "item-button selected" : "item-button"}
                onClick={() => onSelectQueryPath(path.id)}
              >
                <span className="item-title">{path.path_type}</span>
                <span className="item-subtle">{path.id}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="detail-pane">
        <header className="pane-header">
          <h3>Path Detail</h3>
        </header>
        {!selectedPath ? (
          <p className="muted">No query paths in current filter context.</p>
        ) : (
          <dl className="detail-grid">
            <Row label="id" value={<code>{selectedPath.id}</code>} />
            <Row label="path_type" value={selectedPath.path_type} />
            <Row label="entry_record_id" value={<code>{selectedPath.entry_record_id}</code>} />
            <Row label="terminal_record_ids" value={<pre>{JSON.stringify(selectedPath.terminal_record_ids, null, 2)}</pre>} />
            <Row label="steps" value={<pre>{JSON.stringify(selectedPath.steps, null, 2)}</pre>} />
            <Row label="source_evidence" value={<pre>{JSON.stringify(selectedPath.source_evidence, null, 2)}</pre>} />
            <Row label="confidence" value={String(selectedPath.confidence)} />
            <Row
              label="notes"
              value={notesToList(selectedPath.notes).length ? notesToList(selectedPath.notes).join(" | ") : "[]"}
            />
          </dl>
        )}
      </div>
    </section>
  );
}
