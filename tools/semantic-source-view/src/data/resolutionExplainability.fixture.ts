import { ResolutionExplainabilityModel, RESOLUTION_EXPLAINABILITY_SCHEMA_VERSION } from "./resolutionExplainability";

export const PHASE8_RESOLUTION_EXPLAINABILITY_FIXTURE: ResolutionExplainabilityModel = {
  schema_version: RESOLUTION_EXPLAINABILITY_SCHEMA_VERSION,
  mode: "read_only_scaffold",
  record_id: "measure:mlb:air_balls",
  record_type: "measure",
  league: "mlb",
  search_term: "air_balls",
  match_kind: "exact_token",
  match_strength: "medium",
  steps: [
    {
      order: 1,
      kind: "search_entry",
      label: "Search Entry",
      detail: "Fixture example for deterministic explainability rendering.",
      refs: ["record:measure:mlb:air_balls", "search:air_balls"],
    },
    {
      order: 2,
      kind: "semantic_scope",
      label: "Semantic Scope",
      detail: "Record type=measure; league=mlb.",
      refs: ["record_type:measure", "league:mlb"],
    },
    {
      order: 3,
      kind: "source_lineage",
      label: "Source Lineage",
      detail: "Lineage entries=1; evidence entries=1.",
      refs: ["runtime:runtime/leagues/mlb/fetchBaseMeasures.response.json#measures[].air_balls"],
    },
    {
      order: 4,
      kind: "relationship_scope",
      label: "Relationship Scope",
      detail: "Relationship refs=0; related records=0.",
      refs: [],
    },
    {
      order: 5,
      kind: "query_path_scope",
      label: "Query Path Scope",
      detail: "Query path references for this record=0.",
      refs: ["query_paths:0"],
    },
    {
      order: 6,
      kind: "runtime_bridge_deferred",
      label: "Runtime Bridge Deferred",
      detail: "Fixture confirms Phase 8 remains read-only and runtime-bridge deferred.",
      refs: ["deferred:runtime_apply", "deferred:planner_execution", "deferred:runtime_integration"],
    },
  ],
  deferred: [
    "runtime apply behavior",
    "planner execution",
    "SmartStat runtime integration",
    "plan execution",
  ],
};
