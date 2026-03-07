import React, { useMemo } from "react";
import { notesToList } from "../data/viewModel";
import { TypedRecord } from "../types";

interface RecordsPanelProps {
  records: TypedRecord[];
  selectedRecordId: string | null;
  onSelectRecord: (recordId: string) => void;
}

function Row({ label, value }: { label: string; value: React.ReactNode }): JSX.Element {
  return (
    <div className="detail-row">
      <dt>{label}</dt>
      <dd>{value}</dd>
    </div>
  );
}

export function RecordsPanel({
  records,
  selectedRecordId,
  onSelectRecord,
}: RecordsPanelProps): JSX.Element {
  const selectedRecord = useMemo(() => {
    if (!records.length) {
      return null;
    }
    return records.find((record) => record.id === selectedRecordId) ?? records[0];
  }, [records, selectedRecordId]);

  return (
    <section className="split-panel">
      <div className="list-pane">
        <header className="pane-header">
          <h3>Records ({records.length})</h3>
        </header>
        <ul className="item-list">
          {records.map((record) => (
            <li key={record.id}>
              <button
                type="button"
                className={selectedRecord?.id === record.id ? "item-button selected" : "item-button"}
                onClick={() => onSelectRecord(record.id)}
              >
                <span className="item-title">{record.name}</span>
                <span className="item-subtle">{record.id}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="detail-pane">
        <header className="pane-header">
          <h3>Record Detail</h3>
        </header>
        {!selectedRecord ? (
          <p className="muted">No records in current filter context.</p>
        ) : (
          <dl className="detail-grid">
            <Row label="id" value={<code>{selectedRecord.id}</code>} />
            <Row label="name" value={selectedRecord.name} />
            <Row label="record_type" value={selectedRecord.recordType} />
            <Row label="league" value={selectedRecord.league ?? "null"} />
            <Row label="aliases" value={selectedRecord.aliases.length ? selectedRecord.aliases.join(", ") : "[]"} />
            <Row label="notes" value={notesToList(selectedRecord.notes).length ? notesToList(selectedRecord.notes).join(" | ") : "[]"} />
            <Row label="confidence" value={String(selectedRecord.confidence)} />
            <Row label="source_type" value={selectedRecord.source_type} />
            <Row label="source_path" value={<code>{selectedRecord.source_path}</code>} />
            <Row label="source_ref" value={<code>{selectedRecord.source_ref}</code>} />
            <Row label="lineage" value={<pre>{JSON.stringify(selectedRecord.lineage, null, 2)}</pre>} />
            <Row label="evidence" value={<pre>{JSON.stringify(selectedRecord.evidence, null, 2)}</pre>} />
            <Row label="related_ids" value={<pre>{JSON.stringify(selectedRecord.related_ids, null, 2)}</pre>} />
            <Row
              label="relationship_refs"
              value={<pre>{JSON.stringify(selectedRecord.relationship_refs, null, 2)}</pre>}
            />
          </dl>
        )}
      </div>
    </section>
  );
}
