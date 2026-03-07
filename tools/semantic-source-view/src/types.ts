export type SourceType = "runtime" | "lookup_index" | "grammar" | "grammar_snapshot" | "schema";

export type RecordType = "entity" | "measure" | "filter" | "profile" | "qualifier";

export interface LineageEntry {
  role?: string;
  source_type: SourceType;
  source_path: string;
  source_ref: string;
}

export interface SemanticRecord {
  id: string;
  name: string;
  league: string | null;
  source_type: SourceType;
  source_path: string;
  source_ref: string;
  aliases: string[];
  notes?: string[] | string;
  confidence: string | number;
  lineage: LineageEntry[];
  evidence: string[];
  related_ids: string[];
  relationship_refs: string[];
}

export interface RelationshipRecord {
  id: string;
  type: string;
  from_id: string;
  to_id: string;
  league: string | null;
  source_type: SourceType;
  source_path: string;
  source_ref: string;
  confidence: string | number;
  notes?: string[] | string;
}

export interface QueryPathStep {
  kind: string;
  ref_id: string;
  label: string;
  source_path: string | null;
  source_ref: string | null;
}

export interface QueryPathRecord {
  id: string;
  league: string | null;
  entry_record_id: string;
  path_type: string;
  steps: QueryPathStep[];
  terminal_record_ids: string[];
  source_evidence: string[];
  confidence: string | number;
  notes?: string[] | string;
}

export interface TraceRecord {
  lineage: LineageEntry[];
  evidence: string[];
  relationship_ids: string[];
}

export interface TraceIndex {
  by_record_id: Record<string, TraceRecord>;
  by_source_path: Record<string, string[]>;
}

export interface SourceTreeNode {
  id: string;
  label: string;
  node_type: "source_type_root" | "league_bucket" | "misc_bucket" | "source_file";
  source_type: SourceType;
  source_path: string | null;
  league: string | null;
  record_ids: string[];
  children: SourceTreeNode[];
}

export interface SourceTree {
  roots: SourceTreeNode[];
}

export interface UIViewBuckets {
  by_league: Record<string, string[]>;
  by_record_type: Record<string, string[]>;
  by_source_type: Record<string, string[]>;
}

export interface SemanticIndex {
  schema_version: string;
  generated_utc: string;
  source_roots: string[];
  leagues: string[];
  entities: SemanticRecord[];
  measures: SemanticRecord[];
  qualifiers: SemanticRecord[];
  filters: SemanticRecord[];
  profiles: SemanticRecord[];
  relationships: RelationshipRecord[];
  trace_index: TraceIndex;
  query_paths: QueryPathRecord[];
  source_tree: SourceTree;
  ui_views: UIViewBuckets;
  raw_sources: {
    total_files: number;
    roots: Record<
      string,
      {
        exists: boolean;
        file_count: number;
        files: string[];
      }
    >;
  };
}

export interface TypedRecord extends SemanticRecord {
  recordType: RecordType;
}

export interface FilterState {
  league: string;
  recordType: string;
  sourceType: string;
  search: string;
}

export interface TreeSelection {
  nodeId: string;
  label: string;
  sourcePath: string | null;
  sourceType: SourceType | null;
  league: string | null;
  recordIds: string[];
}

export type SearchMatchKind =
  | "browse"
  | "exact_id"
  | "exact_name"
  | "exact_alias"
  | "prefix"
  | "exact_token"
  | "substring"
  | "source_ref"
  | "source_path"
  | "evidence";

export type SearchMatchStrength = "none" | "high" | "medium" | "low";

export interface RecordSearchResult {
  record: TypedRecord;
  match_kind: SearchMatchKind;
  match_strength: SearchMatchStrength;
  match_fields: string[];
  match_scope: "normalized" | "source_hint";
  score: number;
  explanation: string;
}

export interface SearchSourceHint {
  hint_type: "record_source" | "source_tree_label" | "source_path";
  label: string;
  detail: string;
  record_ids: string[];
}

export interface SearchNarrative {
  term: string;
  normalized_match_count: number;
  source_hint_count: number;
  summary: string;
  hints: SearchSourceHint[];
  next_steps: string[];
}
