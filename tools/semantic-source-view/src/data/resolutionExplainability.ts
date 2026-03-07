import { RecordSearchResult } from "../types";

export const RESOLUTION_EXPLAINABILITY_SCHEMA_VERSION = "resolution-explainability.v0";

export type ResolutionExplainabilityStepKind =
  | "search_entry"
  | "semantic_scope"
  | "source_lineage"
  | "relationship_scope"
  | "query_path_scope"
  | "runtime_bridge_deferred";

export interface ResolutionExplainabilityStep {
  order: number;
  kind: ResolutionExplainabilityStepKind;
  label: string;
  detail: string;
  refs: string[];
}

export interface ResolutionExplainabilityModel {
  schema_version: string;
  mode: "read_only_scaffold";
  record_id: string;
  record_type: string;
  league: string | null;
  search_term: string;
  match_kind: string;
  match_strength: string;
  steps: ResolutionExplainabilityStep[];
  deferred: string[];
}

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values)).sort((left, right) => left.localeCompare(right));
}

function normalizeSearchTerm(term: string): string {
  return term.trim();
}

function buildLineageRefs(result: RecordSearchResult): string[] {
  return uniqueSorted(
    result.record.lineage.map(
      (entry) => `${entry.source_type}:${entry.source_path}#${entry.source_ref}`
    )
  );
}

export function buildResolutionExplainabilityModel(
  result: RecordSearchResult,
  searchTerm: string,
  queryPathCount: number
): ResolutionExplainabilityModel {
  const normalizedSearch = normalizeSearchTerm(searchTerm);
  const lineageRefs = buildLineageRefs(result);
  const relationshipRefs = uniqueSorted(result.record.relationship_refs);
  const relatedRecordIds = uniqueSorted(result.record.related_ids);

  const steps: ResolutionExplainabilityStep[] = [
    {
      order: 1,
      kind: "search_entry",
      label: "Search Entry",
      detail: `Match kind=${result.match_kind}; fields=${result.match_fields.join(", ") || "(none)"}.`,
      refs: uniqueSorted([`record:${result.record.id}`, `search:${normalizedSearch || "(browse)"}`]),
    },
    {
      order: 2,
      kind: "semantic_scope",
      label: "Semantic Scope",
      detail: `Record type=${result.record.recordType}; league=${result.record.league ?? "global"}.`,
      refs: uniqueSorted([`record_type:${result.record.recordType}`, `league:${result.record.league ?? "global"}`]),
    },
    {
      order: 3,
      kind: "source_lineage",
      label: "Source Lineage",
      detail: `Lineage entries=${lineageRefs.length}; evidence entries=${result.record.evidence.length}.`,
      refs: uniqueSorted([...lineageRefs, ...result.record.evidence]),
    },
    {
      order: 4,
      kind: "relationship_scope",
      label: "Relationship Scope",
      detail: `Relationship refs=${relationshipRefs.length}; related records=${relatedRecordIds.length}.`,
      refs: uniqueSorted([...relationshipRefs, ...relatedRecordIds]),
    },
    {
      order: 5,
      kind: "query_path_scope",
      label: "Query Path Scope",
      detail: `Query path references for this record=${queryPathCount}.`,
      refs: [`query_paths:${queryPathCount}`],
    },
    {
      order: 6,
      kind: "runtime_bridge_deferred",
      label: "Runtime Bridge Deferred",
      detail:
        "This scaffold is read-only semantic explainability and does not execute runtime/planner behavior.",
      refs: ["deferred:runtime_apply", "deferred:planner_execution", "deferred:runtime_integration"],
    },
  ];

  return {
    schema_version: RESOLUTION_EXPLAINABILITY_SCHEMA_VERSION,
    mode: "read_only_scaffold",
    record_id: result.record.id,
    record_type: result.record.recordType,
    league: result.record.league,
    search_term: normalizedSearch,
    match_kind: result.match_kind,
    match_strength: result.match_strength,
    steps,
    deferred: [
      "runtime apply behavior",
      "planner execution",
      "SmartStat runtime integration",
      "plan execution",
    ],
  };
}
