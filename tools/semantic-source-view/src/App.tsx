import React, { useEffect, useMemo, useState } from "react";
import { FilterBar } from "./components/FilterBar";
import { QueryPathsPanel } from "./components/QueryPathsPanel";
import { RecordsPanel } from "./components/RecordsPanel";
import { RelationshipsPanel } from "./components/RelationshipsPanel";
import { SourceTree } from "./components/SourceTree";
import { TabKey, Tabs } from "./components/Tabs";
import { TraceabilityPanel } from "./components/TraceabilityPanel";
import { loadSemanticIndex } from "./data/loadSemanticIndex";
import {
  buildSearchNarrative,
  buildAllowedRecordIdSet,
  buildFilterOptions,
  buildRecordById,
  filterQueryPathsByContext,
  filterRelationshipsByContext,
  flattenRecords,
  searchRecordsByContext,
} from "./data/viewModel";
import { FilterState, SemanticIndex, TreeSelection } from "./types";

const DEFAULT_FILTERS: FilterState = {
  search: "",
  league: "all",
  recordType: "all",
  sourceType: "all",
};

export default function App(): JSX.Element {
  const [index, setIndex] = useState<SemanticIndex | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<TabKey>("records");
  const [filters, setFilters] = useState<FilterState>(DEFAULT_FILTERS);
  const [selectedNode, setSelectedNode] = useState<TreeSelection | null>(null);
  const [selectedRecordId, setSelectedRecordId] = useState<string | null>(null);
  const [selectedRelationshipId, setSelectedRelationshipId] = useState<string | null>(null);
  const [selectedQueryPathId, setSelectedQueryPathId] = useState<string | null>(null);

  useEffect(() => {
    let mounted = true;
    loadSemanticIndex()
      .then((payload) => {
        if (!mounted) {
          return;
        }
        setIndex(payload);
      })
      .catch((loadError) => {
        if (!mounted) {
          return;
        }
        setError(String(loadError));
      });
    return () => {
      mounted = false;
    };
  }, []);

  const allRecords = useMemo(() => (index ? flattenRecords(index) : []), [index]);
  const recordById = useMemo(() => buildRecordById(allRecords), [allRecords]);
  const filterOptions = useMemo(
    () =>
      index
        ? buildFilterOptions(index)
        : {
            leagues: ["all"],
            recordTypes: ["all"],
            sourceTypes: ["all"],
          },
    [index]
  );

  const allowedRecordIds = useMemo(() => {
    if (!index) {
      return new Set<string>();
    }
    return buildAllowedRecordIdSet(
      index,
      allRecords,
      {
        league: filters.league,
        recordType: filters.recordType,
        sourceType: filters.sourceType,
      },
      selectedNode
    );
  }, [index, allRecords, filters.league, filters.recordType, filters.sourceType, selectedNode]);

  const recordSearch = useMemo(
    () => searchRecordsByContext(allRecords, allowedRecordIds, filters.search),
    [allRecords, allowedRecordIds, filters.search]
  );
  const filteredRecordResults = recordSearch.normalizedResults;
  const sourceOnlyRecordResults = recordSearch.sourceOnlyResults;
  const searchNarrative = useMemo(
    () =>
      index
        ? buildSearchNarrative(
            index,
            filters.search,
            filteredRecordResults,
            sourceOnlyRecordResults,
            allowedRecordIds
          )
        : null,
    [index, filters.search, filteredRecordResults, sourceOnlyRecordResults, allowedRecordIds]
  );

  const filteredRelationships = useMemo(() => {
    if (!index) {
      return [];
    }
    return filterRelationshipsByContext(
      index.relationships,
      allowedRecordIds,
      {
        league: filters.league,
        recordType: filters.recordType,
        sourceType: filters.sourceType,
      },
      selectedNode,
      filters.search
    );
  }, [index, allowedRecordIds, filters.league, filters.recordType, filters.sourceType, selectedNode, filters.search]);

  const filteredQueryPaths = useMemo(() => {
    if (!index) {
      return [];
    }
    return filterQueryPathsByContext(
      index.query_paths,
      recordById,
      allowedRecordIds,
      {
        league: filters.league,
        recordType: filters.recordType,
        sourceType: filters.sourceType,
      },
      selectedNode,
      filters.search
    );
  }, [index, recordById, allowedRecordIds, filters.league, filters.recordType, filters.sourceType, selectedNode, filters.search]);

  useEffect(() => {
    if (
      selectedRecordId &&
      !filteredRecordResults.some((result) => result.record.id === selectedRecordId)
    ) {
      setSelectedRecordId(null);
    }
  }, [filteredRecordResults, selectedRecordId]);

  useEffect(() => {
    if (
      selectedRelationshipId &&
      !filteredRelationships.some((relationship) => relationship.id === selectedRelationshipId)
    ) {
      setSelectedRelationshipId(null);
    }
  }, [filteredRelationships, selectedRelationshipId]);

  useEffect(() => {
    if (selectedQueryPathId && !filteredQueryPaths.some((path) => path.id === selectedQueryPathId)) {
      setSelectedQueryPathId(null);
    }
  }, [filteredQueryPaths, selectedQueryPathId]);

  if (error) {
    return (
      <div className="app-shell">
        <main className="error-box">
          <h1>SmartStat Semantic Source View</h1>
          <p>Failed to load semantic index.</p>
          <pre>{error}</pre>
        </main>
      </div>
    );
  }

  if (!index) {
    return (
      <div className="app-shell">
        <main className="loading-box">
          <h1>SmartStat Semantic Source View</h1>
          <p>Loading `.tools/onair_dump/index/semantic_index.json`...</p>
        </main>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>SmartStat Semantic Source View</h1>
        <p className="subtitle">
          Read-only semantic inspection UI for deterministic records, relationships, query paths, and
          traceability.
        </p>
      </header>

      <FilterBar
        filters={filters}
        leagues={filterOptions.leagues}
        recordTypes={filterOptions.recordTypes}
        sourceTypes={filterOptions.sourceTypes}
        onChange={setFilters}
      />

      <section className="workspace">
        <SourceTree
          roots={index.source_tree.roots}
          selectedNodeId={selectedNode?.nodeId ?? null}
          onSelect={setSelectedNode}
        />

        <main className="main-pane">
          <div className="context-strip">
            <span>Selected Node:</span>
            <strong>{selectedNode ? selectedNode.label : "All"}</strong>
          </div>

          <Tabs value={activeTab} onChange={setActiveTab} />

          {activeTab === "records" ? (
            <RecordsPanel
              results={filteredRecordResults}
              sourceOnlyResults={sourceOnlyRecordResults}
              searchTerm={filters.search}
              narrative={searchNarrative}
              queryPaths={index.query_paths}
              selectedRecordId={selectedRecordId}
              onSelectRecord={setSelectedRecordId}
            />
          ) : null}

          {activeTab === "relationships" ? (
            <RelationshipsPanel
              relationships={filteredRelationships}
              selectedRelationshipId={selectedRelationshipId}
              onSelectRelationship={setSelectedRelationshipId}
            />
          ) : null}

          {activeTab === "query_paths" ? (
            <QueryPathsPanel
              queryPaths={filteredQueryPaths}
              selectedQueryPathId={selectedQueryPathId}
              onSelectQueryPath={setSelectedQueryPathId}
            />
          ) : null}

          {activeTab === "traceability" ? (
            <TraceabilityPanel
              traceIndex={index.trace_index}
              allowedRecordIds={allowedRecordIds}
              sourceTypeFilter={filters.sourceType}
              search={filters.search}
              selectedNode={selectedNode}
            />
          ) : null}
        </main>
      </section>
    </div>
  );
}
