import React, { useMemo } from "react";
import { matchLabel, notesToList } from "../data/viewModel";
import { QueryPathRecord, RecordSearchResult, SearchNarrative } from "../types";

interface RecordsPanelProps {
  results: RecordSearchResult[];
  sourceOnlyResults: RecordSearchResult[];
  searchTerm: string;
  narrative: SearchNarrative | null;
  queryPaths: QueryPathRecord[];
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

function formatLeague(league: string | null): string {
  return league ? league.toUpperCase() : "global";
}

function buildExistenceNotes(result: RecordSearchResult): string[] {
  const lineage = result.record.lineage;
  if (!lineage.length) {
    return ["No lineage metadata available."];
  }

  const notes: string[] = [];
  const primary = lineage[0];
  notes.push(
    `${primary.source_type} source introduced this record (${primary.source_path}#${primary.source_ref}).`
  );
  if (lineage.some((entry) => entry.source_type === "lookup_index")) {
    notes.push("lookup_index lineage indicates enrichment/alias context.");
  }
  if (lineage.some((entry) => entry.source_type === "runtime")) {
    notes.push("runtime lineage indicates data surfaced from runtime payloads.");
  }
  return notes;
}

function EmptySearchState({
  narrative,
  sourceOnlyResults,
}: {
  narrative: SearchNarrative | null;
  sourceOnlyResults: RecordSearchResult[];
}): JSX.Element {
  return (
    <div className="empty-state">
      <h4>No normalized records matched</h4>
      <p>{narrative?.summary ?? "No normalized records matched this term."}</p>
      {narrative?.hints.length ? (
        <>
          <h5>Source-oriented hints</h5>
          <ul>
            {narrative.hints.map((hint) => (
              <li key={`${hint.hint_type}:${hint.label}:${hint.detail}`}>
                <strong>{hint.hint_type}</strong> - <code>{hint.label}</code> - {hint.detail}
              </li>
            ))}
          </ul>
        </>
      ) : null}
      {!narrative?.hints.length && sourceOnlyResults.length ? (
        <>
          <h5>Record source hints</h5>
          <ul>
            {sourceOnlyResults.slice(0, 8).map((hint) => (
              <li key={hint.record.id}>
                <code>{hint.record.id}</code> - {hint.explanation}
              </li>
            ))}
          </ul>
        </>
      ) : null}
      <h5>Try next</h5>
      <ul>
        {(narrative?.next_steps ?? []).map((step) => (
          <li key={step}>{step}</li>
        ))}
      </ul>
    </div>
  );
}

export function RecordsPanel({
  results,
  sourceOnlyResults,
  searchTerm,
  narrative,
  queryPaths,
  selectedRecordId,
  onSelectRecord,
}: RecordsPanelProps): JSX.Element {
  const selectedResult = useMemo(() => {
    if (!results.length) {
      return null;
    }
    return results.find((result) => result.record.id === selectedRecordId) ?? results[0];
  }, [results, selectedRecordId]);

  const relatedMatches = useMemo(() => {
    if (!searchTerm.trim() || !selectedResult) {
      return [];
    }
    return results
      .filter((result) => result.record.id !== selectedResult.record.id)
      .slice(0, 3);
  }, [searchTerm, results, selectedResult]);

  const queryPathCountForRecord = useMemo(() => {
    if (!selectedResult) {
      return 0;
    }
    const id = selectedResult.record.id;
    return queryPaths.filter(
      (path) => path.entry_record_id === id || path.terminal_record_ids.includes(id)
    ).length;
  }, [selectedResult, queryPaths]);

  return (
    <section className="split-panel">
      <div className="list-pane">
        <header className="pane-header">
          <h3>Records ({results.length})</h3>
        </header>
        <ul className="item-list">
          {results.map((result) => (
            <li key={result.record.id}>
              <button
                type="button"
                className={
                  selectedResult?.record.id === result.record.id ? "item-button selected" : "item-button"
                }
                onClick={() => onSelectRecord(result.record.id)}
              >
                <span className="item-title">{result.record.name}</span>
                <span className="badge-row">
                  <span className="badge">{result.record.recordType}</span>
                  <span className="badge">{formatLeague(result.record.league)}</span>
                  <span className="badge">{result.match_strength}</span>
                  <span className="badge">{matchLabel(result.match_kind)}</span>
                </span>
                <span className="item-subtle">{result.record.id}</span>
                <span className="item-snippet">{result.explanation}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div className="detail-pane">
        <header className="pane-header">
          <h3>Record Detail</h3>
        </header>
        {!selectedResult ? (
          <EmptySearchState narrative={narrative} sourceOnlyResults={sourceOnlyResults} />
        ) : (
          <>
            <section className="narrative-panel">
              <h4>Debug Narrative</h4>
              <p>
                This is a <strong>{selectedResult.record.recordType}</strong> record for league{" "}
                <strong>{formatLeague(selectedResult.record.league)}</strong>.
              </p>
              <p>
                Search explanation: {selectedResult.explanation}. Match fields:{" "}
                {selectedResult.match_fields.length
                  ? selectedResult.match_fields.join(", ")
                  : "(none)"}.
              </p>
              <ul>
                {buildExistenceNotes(selectedResult).map((note) => (
                  <li key={note}>{note}</li>
                ))}
              </ul>
              <p>
                Next places to inspect: Relationships ({selectedResult.record.relationship_refs.length}),
                Query Paths ({queryPathCountForRecord}), and Traceability lineage.
              </p>
              {relatedMatches.length ? (
                <>
                  <h5>Related matches</h5>
                  <ul>
                    {relatedMatches.map((result) => (
                      <li key={result.record.id}>
                        <code>{result.record.id}</code> - {result.explanation}
                      </li>
                    ))}
                  </ul>
                </>
              ) : null}
            </section>

            <dl className="detail-grid">
              <Row label="id" value={<code>{selectedResult.record.id}</code>} />
              <Row label="name" value={selectedResult.record.name} />
              <Row label="record_type" value={selectedResult.record.recordType} />
              <Row label="league" value={selectedResult.record.league ?? "null"} />
              <Row
                label="aliases"
                value={
                  selectedResult.record.aliases.length
                    ? selectedResult.record.aliases.join(", ")
                    : "[]"
                }
              />
              <Row
                label="notes"
                value={
                  notesToList(selectedResult.record.notes).length
                    ? notesToList(selectedResult.record.notes).join(" | ")
                    : "[]"
                }
              />
              <Row label="confidence" value={String(selectedResult.record.confidence)} />
              <Row label="source_type" value={selectedResult.record.source_type} />
              <Row label="source_path" value={<code>{selectedResult.record.source_path}</code>} />
              <Row label="source_ref" value={<code>{selectedResult.record.source_ref}</code>} />
              <Row
                label="lineage"
                value={<pre>{JSON.stringify(selectedResult.record.lineage, null, 2)}</pre>}
              />
              <Row
                label="evidence"
                value={<pre>{JSON.stringify(selectedResult.record.evidence, null, 2)}</pre>}
              />
              <Row
                label="related_ids"
                value={<pre>{JSON.stringify(selectedResult.record.related_ids, null, 2)}</pre>}
              />
              <Row
                label="relationship_refs"
                value={<pre>{JSON.stringify(selectedResult.record.relationship_refs, null, 2)}</pre>}
              />
            </dl>
          </>
        )}
      </div>
    </section>
  );
}
