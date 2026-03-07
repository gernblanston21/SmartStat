import React, { useMemo } from "react";
import { notesToList } from "../data/viewModel";
import { RelationshipRecord } from "../types";

interface RelationshipsPanelProps {
  relationships: RelationshipRecord[];
  selectedRelationshipId: string | null;
  onSelectRelationship: (relationshipId: string) => void;
}

function Row({ label, value }: { label: string; value: React.ReactNode }): JSX.Element {
  return (
    <div className="detail-row">
      <dt>{label}</dt>
      <dd>{value}</dd>
    </div>
  );
}

export function RelationshipsPanel({
  relationships,
  selectedRelationshipId,
  onSelectRelationship,
}: RelationshipsPanelProps): JSX.Element {
  const selectedRelationship = useMemo(() => {
    if (!relationships.length) {
      return null;
    }
    return relationships.find((item) => item.id === selectedRelationshipId) ?? relationships[0];
  }, [relationships, selectedRelationshipId]);

  return (
    <section className="split-panel">
      <div className="list-pane">
        <header className="pane-header">
          <h3>Relationships ({relationships.length})</h3>
        </header>
        <ul className="item-list">
          {relationships.map((relationship) => (
            <li key={relationship.id}>
              <button
                type="button"
                className={
                  selectedRelationship?.id === relationship.id ? "item-button selected" : "item-button"
                }
                onClick={() => onSelectRelationship(relationship.id)}
              >
                <span className="item-title">{relationship.type}</span>
                <span className="item-subtle">{relationship.id}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="detail-pane">
        <header className="pane-header">
          <h3>Relationship Detail</h3>
        </header>
        {!selectedRelationship ? (
          <p className="muted">No relationships in current filter context.</p>
        ) : (
          <dl className="detail-grid">
            <Row label="id" value={<code>{selectedRelationship.id}</code>} />
            <Row label="type" value={selectedRelationship.type} />
            <Row label="from_id" value={<code>{selectedRelationship.from_id}</code>} />
            <Row label="to_id" value={<code>{selectedRelationship.to_id}</code>} />
            <Row label="league" value={selectedRelationship.league ?? "null"} />
            <Row label="source_type" value={selectedRelationship.source_type} />
            <Row label="source_path" value={<code>{selectedRelationship.source_path}</code>} />
            <Row label="source_ref" value={<code>{selectedRelationship.source_ref}</code>} />
            <Row label="confidence" value={String(selectedRelationship.confidence)} />
            <Row
              label="notes"
              value={notesToList(selectedRelationship.notes).length ? notesToList(selectedRelationship.notes).join(" | ") : "[]"}
            />
          </dl>
        )}
      </div>
    </section>
  );
}
