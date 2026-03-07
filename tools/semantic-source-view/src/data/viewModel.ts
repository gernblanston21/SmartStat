import {
  FilterState,
  QueryPathRecord,
  RecordSearchResult,
  RecordType,
  RelationshipRecord,
  SearchMatchKind,
  SearchNarrative,
  SearchSourceHint,
  SemanticIndex,
  TreeSelection,
  TypedRecord,
} from "../types";
import {
  SEARCH_MATCH_SCORE,
  SOURCE_ONLY_MATCH_KINDS,
  compareRecordSearchResults,
  matchStrengthForKind,
  normalizeSearchText,
  splitSearchTokens,
} from "./searchContract";

const RECORD_TYPE_TO_UI_KEY: Record<RecordType, string> = {
  entity: "entities",
  measure: "measures",
  filter: "filters",
  profile: "profiles",
  qualifier: "qualifiers",
};

const UI_KEY_TO_RECORD_TYPE: Record<string, RecordType> = {
  entities: "entity",
  measures: "measure",
  filters: "filter",
  profiles: "profile",
  qualifiers: "qualifier",
};

function setFromList(list: string[] | undefined): Set<string> {
  return new Set(list ?? []);
}

function intersect(base: Set<string>, candidate: Set<string>): Set<string> {
  const result = new Set<string>();
  for (const value of base) {
    if (candidate.has(value)) {
      result.add(value);
    }
  }
  return result;
}

function makeSearchResult(
  record: TypedRecord,
  kind: SearchMatchKind,
  fields: string[],
  explanation: string
): RecordSearchResult {
  return {
    record,
    match_kind: kind,
    match_strength: matchStrengthForKind(kind),
    match_fields: fields,
    match_scope: SOURCE_ONLY_MATCH_KINDS.has(kind) ? "source_hint" : "normalized",
    score: SEARCH_MATCH_SCORE[kind],
    explanation,
  };
}

function evaluateRecordMatch(record: TypedRecord, rawQuery: string): RecordSearchResult | null {
  const qLower = rawQuery.trim().toLowerCase();
  const qNorm = normalizeSearchText(rawQuery);

  if (!qNorm) {
    return makeSearchResult(record, "browse", [], "Browse mode (no search term)");
  }

  const idLower = record.id.toLowerCase();
  const nameLower = record.name.toLowerCase();
  const aliasesLower = record.aliases.map((x) => x.toLowerCase());
  const sourceRefLower = record.source_ref.toLowerCase();
  const sourcePathLower = record.source_path.toLowerCase();
  const evidenceLower = record.evidence.map((x) => x.toLowerCase());

  const idNorm = normalizeSearchText(record.id);
  const idTailNorm = normalizeSearchText(record.id.split(":").slice(2).join(":"));
  const nameNorm = normalizeSearchText(record.name);
  const aliasesNorm = record.aliases.map((x) => normalizeSearchText(x));
  const sourceRefNorm = normalizeSearchText(record.source_ref);
  const sourcePathNorm = normalizeSearchText(record.source_path);
  const evidenceNorm = record.evidence.map((x) => normalizeSearchText(x));

  if (idLower === qLower || idNorm === qNorm) {
    return makeSearchResult(record, "exact_id", ["id"], "Matched exact ID");
  }

  if (nameLower === qLower || nameNorm === qNorm) {
    return makeSearchResult(record, "exact_name", ["name"], "Matched exact name");
  }

  if (aliasesLower.includes(qLower) || aliasesNorm.includes(qNorm)) {
    return makeSearchResult(record, "exact_alias", ["aliases"], "Matched exact alias");
  }

  const prefixFields: string[] = [];
  if (idNorm.startsWith(qNorm) || idTailNorm.startsWith(qNorm)) {
    prefixFields.push("id");
  }
  if (nameNorm.startsWith(qNorm)) {
    prefixFields.push("name");
  }
  if (aliasesNorm.some((x) => x.startsWith(qNorm))) {
    prefixFields.push("aliases");
  }
  if (prefixFields.length) {
    return makeSearchResult(record, "prefix", prefixFields, "Matched prefix in ID/name/alias");
  }

  const tokenSet = new Set<string>([
    ...splitSearchTokens(idNorm),
    ...splitSearchTokens(idTailNorm),
    ...splitSearchTokens(nameNorm),
    ...aliasesNorm.flatMap((x) => splitSearchTokens(x)),
  ]);
  if (tokenSet.has(qNorm)) {
    return makeSearchResult(record, "exact_token", ["id", "name"], "Matched exact normalized token");
  }

  const substringFields: string[] = [];
  if (idNorm.includes(qNorm) || idTailNorm.includes(qNorm)) {
    substringFields.push("id");
  }
  if (nameNorm.includes(qNorm)) {
    substringFields.push("name");
  }
  if (aliasesNorm.some((x) => x.includes(qNorm))) {
    substringFields.push("aliases");
  }
  if (substringFields.length) {
    return makeSearchResult(record, "substring", substringFields, "Matched partial ID/name/alias");
  }

  if (sourceRefNorm.includes(qNorm) || sourceRefLower.includes(qLower)) {
    return makeSearchResult(record, "source_ref", ["source_ref"], "Matched source reference");
  }

  if (sourcePathNorm.includes(qNorm) || sourcePathLower.includes(qLower)) {
    return makeSearchResult(record, "source_path", ["source_path"], "Matched source path");
  }

  if (evidenceNorm.some((x) => x.includes(qNorm)) || evidenceLower.some((x) => x.includes(qLower))) {
    return makeSearchResult(record, "evidence", ["evidence"], "Matched source evidence");
  }

  return null;
}

function collectSourceTreeHints(
  index: SemanticIndex,
  qNorm: string,
  allowedRecordIds: Set<string>
): SearchSourceHint[] {
  const hints: SearchSourceHint[] = [];
  const stack = [...index.source_tree.roots];
  while (stack.length) {
    const node = stack.shift()!;
    stack.push(...node.children);

    const labelNorm = normalizeSearchText(node.label);
    const pathNorm = normalizeSearchText(node.source_path ?? "");
    if (!labelNorm.includes(qNorm) && !pathNorm.includes(qNorm)) {
      continue;
    }

    const matchedRecordIds = node.record_ids.filter((id) => allowedRecordIds.has(id));
    hints.push({
      hint_type: "source_tree_label",
      label: node.label,
      detail: node.source_path ?? node.node_type,
      record_ids: matchedRecordIds,
    });
  }
  return hints;
}

function collectTracePathHints(
  index: SemanticIndex,
  rawQuery: string,
  qNorm: string,
  allowedRecordIds: Set<string>
): SearchSourceHint[] {
  const qLower = rawQuery.trim().toLowerCase();
  const hints: SearchSourceHint[] = [];
  for (const sourcePath of Object.keys(index.trace_index.by_source_path)) {
    const pathNorm = normalizeSearchText(sourcePath);
    if (!pathNorm.includes(qNorm) && !sourcePath.toLowerCase().includes(qLower)) {
      continue;
    }
    const recordIds = (index.trace_index.by_source_path[sourcePath] ?? []).filter((id) =>
      allowedRecordIds.has(id)
    );
    if (!recordIds.length) {
      continue;
    }
    hints.push({
      hint_type: "source_path",
      label: sourcePath,
      detail: "Matched trace source path",
      record_ids: recordIds,
    });
  }
  return hints;
}

function compareHints(left: SearchSourceHint, right: SearchSourceHint): number {
  const typeCmp = left.hint_type.localeCompare(right.hint_type);
  if (typeCmp !== 0) {
    return typeCmp;
  }
  return left.label.localeCompare(right.label);
}

export function flattenRecords(index: SemanticIndex): TypedRecord[] {
  return [
    ...index.entities.map((record) => ({ ...record, recordType: "entity" as const })),
    ...index.measures.map((record) => ({ ...record, recordType: "measure" as const })),
    ...index.filters.map((record) => ({ ...record, recordType: "filter" as const })),
    ...index.profiles.map((record) => ({ ...record, recordType: "profile" as const })),
    ...index.qualifiers.map((record) => ({ ...record, recordType: "qualifier" as const })),
  ];
}

export function buildRecordById(records: TypedRecord[]): Record<string, TypedRecord> {
  return records.reduce<Record<string, TypedRecord>>((acc, record) => {
    acc[record.id] = record;
    return acc;
  }, {});
}

export function buildFilterOptions(index: SemanticIndex): {
  leagues: string[];
  recordTypes: string[];
  sourceTypes: string[];
} {
  const leagueOptions = ["all", ...Object.keys(index.ui_views.by_league)];
  const recordTypeOptions = [
    "all",
    ...Object.keys(index.ui_views.by_record_type)
      .map((key) => UI_KEY_TO_RECORD_TYPE[key])
      .filter((value): value is RecordType => Boolean(value)),
  ];
  const sourceTypeOptions = ["all", ...Object.keys(index.ui_views.by_source_type)];

  return {
    leagues: leagueOptions,
    recordTypes: recordTypeOptions,
    sourceTypes: sourceTypeOptions,
  };
}

export function buildAllowedRecordIdSet(
  index: SemanticIndex,
  allRecords: TypedRecord[],
  filters: Omit<FilterState, "search">,
  selectedNode: TreeSelection | null
): Set<string> {
  let allowed = new Set<string>(allRecords.map((record) => record.id));

  if (filters.league !== "all") {
    allowed = intersect(allowed, setFromList(index.ui_views.by_league[filters.league]));
  }

  if (filters.recordType !== "all") {
    const uiKey = RECORD_TYPE_TO_UI_KEY[filters.recordType as RecordType];
    if (uiKey) {
      allowed = intersect(allowed, setFromList(index.ui_views.by_record_type[uiKey]));
    }
  }

  if (filters.sourceType !== "all") {
    allowed = intersect(allowed, setFromList(index.ui_views.by_source_type[filters.sourceType]));
  }

  if (selectedNode) {
    allowed = intersect(allowed, new Set<string>(selectedNode.recordIds));
  }

  return allowed;
}

export function searchRecordsByContext(
  records: TypedRecord[],
  allowedRecordIds: Set<string>,
  search: string
): {
  normalizedResults: RecordSearchResult[];
  sourceOnlyResults: RecordSearchResult[];
} {
  const matches: RecordSearchResult[] = [];
  for (const record of records) {
    if (!allowedRecordIds.has(record.id)) {
      continue;
    }
    const evaluated = evaluateRecordMatch(record, search);
    if (evaluated) {
      matches.push(evaluated);
    }
  }

  matches.sort(compareRecordSearchResults);
  return {
    normalizedResults: matches.filter((x) => x.match_scope === "normalized"),
    sourceOnlyResults: matches.filter((x) => x.match_scope === "source_hint"),
  };
}

export function buildSearchNarrative(
  index: SemanticIndex,
  search: string,
  normalizedResults: RecordSearchResult[],
  sourceOnlyResults: RecordSearchResult[],
  allowedRecordIds: Set<string>
): SearchNarrative | null {
  const raw = search.trim();
  if (!raw) {
    return null;
  }

  const qNorm = normalizeSearchText(raw);
  if (!qNorm) {
    return null;
  }

  const sourceHints: SearchSourceHint[] = [];
  for (const result of sourceOnlyResults.slice(0, 12)) {
    sourceHints.push({
      hint_type: "record_source",
      label: result.record.id,
      detail: result.explanation,
      record_ids: [result.record.id],
    });
  }
  sourceHints.push(...collectSourceTreeHints(index, qNorm, allowedRecordIds));
  sourceHints.push(...collectTracePathHints(index, raw, qNorm, allowedRecordIds));

  const dedup = new Map<string, SearchSourceHint>();
  for (const hint of sourceHints) {
    const key = `${hint.hint_type}|${hint.label}|${hint.detail}`;
    if (!dedup.has(key)) {
      dedup.set(key, hint);
    }
  }
  const uniqueHints = Array.from(dedup.values()).sort(compareHints).slice(0, 14);

  const normalizedCount = normalizedResults.length;
  const sourceHintCount = uniqueHints.length;
  let summary = "";
  if (normalizedCount > 0) {
    summary = `${normalizedCount} normalized record(s) matched this term. Results are ranked by deterministic match strength.`;
  } else {
    summary = "No normalized records matched this term.";
    if (sourceHintCount > 0) {
      summary += " The term still appears in source-oriented data hints.";
    } else {
      summary += " No source hints were found in the current filter context.";
    }
  }

  return {
    term: raw,
    normalized_match_count: normalizedCount,
    source_hint_count: sourceHintCount,
    summary,
    hints: uniqueHints,
    next_steps: [
      "Inspect Source Tree for lookup/runtime paths.",
      "Open Traceability to follow source-path links.",
      "Check Relationships and Query Paths for nearby semantic context.",
    ],
  };
}

export function filterRelationshipsByContext(
  relationships: RelationshipRecord[],
  allowedRecordIds: Set<string>,
  filters: Omit<FilterState, "search">,
  selectedNode: TreeSelection | null,
  search: string
): RelationshipRecord[] {
  const q = search.trim().toLowerCase();
  return relationships.filter((relationship) => {
    if (filters.league !== "all" && relationship.league !== filters.league) {
      return false;
    }
    if (filters.sourceType !== "all" && relationship.source_type !== filters.sourceType) {
      return false;
    }

    const relatesToAllowed =
      allowedRecordIds.has(relationship.from_id) || allowedRecordIds.has(relationship.to_id);
    const sourceMatch = selectedNode?.sourcePath && relationship.source_path === selectedNode.sourcePath;
    if (selectedNode && !relatesToAllowed && !sourceMatch) {
      return false;
    }
    if (!selectedNode && !relatesToAllowed) {
      return false;
    }

    if (!q) {
      return true;
    }
    return (
      relationship.id.toLowerCase().includes(q) ||
      relationship.type.toLowerCase().includes(q) ||
      relationship.from_id.toLowerCase().includes(q) ||
      relationship.to_id.toLowerCase().includes(q) ||
      relationship.source_path.toLowerCase().includes(q) ||
      relationship.source_ref.toLowerCase().includes(q)
    );
  });
}

export function filterQueryPathsByContext(
  queryPaths: QueryPathRecord[],
  recordById: Record<string, TypedRecord>,
  allowedRecordIds: Set<string>,
  filters: Omit<FilterState, "search">,
  selectedNode: TreeSelection | null,
  search: string
): QueryPathRecord[] {
  const q = search.trim().toLowerCase();
  return queryPaths.filter((path) => {
    if (filters.league !== "all" && path.league !== filters.league) {
      return false;
    }

    const entryRecord = recordById[path.entry_record_id];
    if (filters.recordType !== "all" && entryRecord?.recordType !== filters.recordType) {
      return false;
    }
    if (filters.sourceType !== "all" && entryRecord?.source_type !== filters.sourceType) {
      return false;
    }

    const contextMatch =
      allowedRecordIds.has(path.entry_record_id) ||
      path.terminal_record_ids.some((id) => allowedRecordIds.has(id)) ||
      (selectedNode?.sourcePath && path.steps.some((step) => step.source_path === selectedNode.sourcePath));
    if (!contextMatch) {
      return false;
    }

    if (!q) {
      return true;
    }
    return (
      path.id.toLowerCase().includes(q) ||
      path.path_type.toLowerCase().includes(q) ||
      path.entry_record_id.toLowerCase().includes(q)
    );
  });
}

export function notesToList(notes: string[] | string | undefined): string[] {
  if (!notes) {
    return [];
  }
  if (Array.isArray(notes)) {
    return notes;
  }
  return [notes];
}

export { matchLabel, matchStrengthForKind } from "./searchContract";
