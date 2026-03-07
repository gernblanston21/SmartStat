import {
  FilterState,
  QueryPathRecord,
  RecordType,
  RelationshipRecord,
  SemanticIndex,
  TreeSelection,
  TypedRecord,
} from "../types";

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

export function filterRecordsByContext(
  records: TypedRecord[],
  allowedRecordIds: Set<string>,
  search: string
): TypedRecord[] {
  const q = search.trim().toLowerCase();
  return records.filter((record) => {
    if (!allowedRecordIds.has(record.id)) {
      return false;
    }
    if (!q) {
      return true;
    }
    return (
      record.id.toLowerCase().includes(q) ||
      record.name.toLowerCase().includes(q) ||
      (record.league ?? "").toLowerCase().includes(q) ||
      record.source_path.toLowerCase().includes(q)
    );
  });
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
    const sourceMatch =
      selectedNode?.sourcePath && relationship.source_path === selectedNode.sourcePath;
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
      (selectedNode?.sourcePath &&
        path.steps.some((step) => step.source_path === selectedNode.sourcePath));
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
