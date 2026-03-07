import { describe, expect, it } from "vitest";
import { SemanticIndex, SemanticRecord, TypedRecord } from "../types";
import { buildSearchNarrative, searchRecordsByContext } from "./viewModel";

function makeRecord(
  id: string,
  name: string,
  overrides: Partial<TypedRecord> = {}
): TypedRecord {
  const sourceType = overrides.source_type ?? "runtime";
  const sourcePath =
    overrides.source_path ?? "runtime/leagues/mlb/fetchBaseMeasures.response.json";
  const sourceRef = overrides.source_ref ?? "measures[]";

  return {
    id,
    name,
    league: overrides.league ?? "mlb",
    source_type: sourceType,
    source_path: sourcePath,
    source_ref: sourceRef,
    aliases: overrides.aliases ?? [],
    notes: overrides.notes,
    confidence: overrides.confidence ?? "high",
    lineage:
      overrides.lineage ?? [
        {
          source_type: sourceType,
          source_path: sourcePath,
          source_ref: sourceRef,
        },
      ],
    evidence: overrides.evidence ?? [],
    related_ids: overrides.related_ids ?? [],
    relationship_refs: overrides.relationship_refs ?? [],
    recordType: overrides.recordType ?? "measure",
  };
}

function toSemanticRecords(records: TypedRecord[], recordType: TypedRecord["recordType"]): SemanticRecord[] {
  return records
    .filter((record) => record.recordType === recordType)
    .map(({ recordType: _recordType, ...semanticRecord }) => semanticRecord);
}

function buildSemanticIndexFixture(records: TypedRecord[]): SemanticIndex {
  const recordIds = records.map((record) => record.id).sort();
  const firstRecordId = recordIds[0] ?? "";
  const traceSourcePath = "lookup_index/mlb/playerSplits.lookup_index.json";

  const byRecordType: Record<string, string[]> = {
    entities: [],
    measures: [],
    filters: [],
    profiles: [],
    qualifiers: [],
  };
  for (const record of records) {
    const key = `${record.recordType}s`;
    byRecordType[key] = [...(byRecordType[key] ?? []), record.id].sort();
  }

  const byLeague: Record<string, string[]> = {};
  for (const record of records) {
    const leagueKey = record.league ?? "global";
    byLeague[leagueKey] = [...(byLeague[leagueKey] ?? []), record.id].sort();
  }

  return {
    schema_version: "semantic-index.v1",
    generated_utc: "1970-01-01T00:00:00Z",
    source_roots: ["runtime", "lookup_index"],
    leagues: ["mlb"],
    entities: toSemanticRecords(records, "entity"),
    measures: toSemanticRecords(records, "measure"),
    qualifiers: toSemanticRecords(records, "qualifier"),
    filters: toSemanticRecords(records, "filter"),
    profiles: toSemanticRecords(records, "profile"),
    relationships: [],
    trace_index: {
      by_record_id: Object.fromEntries(
        records.map((record) => [
          record.id,
          {
            lineage: record.lineage,
            evidence: record.evidence,
            relationship_ids: [],
          },
        ])
      ),
      by_source_path: {
        [traceSourcePath]: firstRecordId ? [firstRecordId] : [],
      },
    },
    query_paths: [],
    source_tree: {
      roots: [
        {
          id: "root:lookup_index",
          label: "lookup_index",
          node_type: "source_type_root",
          source_type: "lookup_index",
          source_path: null,
          league: null,
          record_ids: recordIds,
          children: [
            {
              id: "source:playerSplits",
              label: "playerSplits bucket",
              node_type: "source_file",
              source_type: "lookup_index",
              source_path: traceSourcePath,
              league: "mlb",
              record_ids: firstRecordId ? [firstRecordId] : [],
              children: [],
            },
          ],
        },
      ],
    },
    ui_views: {
      by_league: byLeague,
      by_record_type: byRecordType,
      by_source_type: {
        runtime: records
          .filter((record) => record.source_type === "runtime")
          .map((record) => record.id)
          .sort(),
        lookup_index: records
          .filter((record) => record.source_type === "lookup_index")
          .map((record) => record.id)
          .sort(),
        grammar: [],
        grammar_snapshot: [],
        schema: [],
      },
    },
    raw_sources: {
      total_files: 0,
      roots: {},
    },
  };
}

describe("deterministic search ranking contract", () => {
  it("exact ID outranks exact name, alias, prefix, and substring matches", () => {
    const records: TypedRecord[] = [
      makeRecord("air_balls", "Air Balls Canonical"),
      makeRecord("measure:mlb:name_match", "Air Balls"),
      makeRecord("measure:mlb:alias_match", "Alias Match", { aliases: ["air_balls"] }),
      makeRecord("measure:mlb:air_balls_percentage", "Air Balls Percentage"),
      makeRecord("measure:mlb:total_air_balls_rate", "Total Air Balls Rate"),
    ];
    const allowed = new Set(records.map((record) => record.id));

    const { normalizedResults } = searchRecordsByContext(records, allowed, "air_balls");
    const rankedIds = normalizedResults.map((result) => result.record.id);

    const exactIdIndex = rankedIds.indexOf("air_balls");
    expect(exactIdIndex).toBe(0);
    expect(exactIdIndex).toBeLessThan(rankedIds.indexOf("measure:mlb:name_match"));
    expect(exactIdIndex).toBeLessThan(rankedIds.indexOf("measure:mlb:alias_match"));
    expect(exactIdIndex).toBeLessThan(rankedIds.indexOf("measure:mlb:air_balls_percentage"));
    expect(exactIdIndex).toBeLessThan(rankedIds.indexOf("measure:mlb:total_air_balls_rate"));
  });

  it("ranks measure:mlb:air_balls above measure:mlb:air_balls_percentage for air_balls", () => {
    const records: TypedRecord[] = [
      makeRecord("measure:mlb:air_balls", "Air Balls"),
      makeRecord("measure:mlb:air_balls_percentage", "Air Balls Percentage"),
    ];
    const allowed = new Set(records.map((record) => record.id));

    const { normalizedResults } = searchRecordsByContext(records, allowed, "air_balls");
    expect(normalizedResults.map((result) => result.record.id)).toEqual([
      "measure:mlb:air_balls",
      "measure:mlb:air_balls_percentage",
    ]);
  });

  it("keeps source-only matches out of normalized results", () => {
    const records: TypedRecord[] = [
      makeRecord("measure:mlb:strikeouts", "Strikeouts", {
        source_type: "lookup_index",
        source_ref: "lookup.playerSplits.bucket",
      }),
    ];
    const allowed = new Set(records.map((record) => record.id));

    const results = searchRecordsByContext(records, allowed, "playerSplits");

    expect(results.normalizedResults).toHaveLength(0);
    expect(results.sourceOnlyResults).toHaveLength(1);
    expect(results.sourceOnlyResults[0].record.id).toBe("measure:mlb:strikeouts");
    expect(results.sourceOnlyResults[0].match_scope).toBe("source_hint");
  });

  it("returns browse-mode normalized results in deterministic order for empty search", () => {
    const records: TypedRecord[] = [
      makeRecord("measure:mlb:m1", "M1", { recordType: "measure", league: "mlb" }),
      makeRecord("entity:nba:e1", "E1", { recordType: "entity", league: "nba" }),
      makeRecord("measure:nba:m0", "M0", { recordType: "measure", league: "nba" }),
      makeRecord("filter:global:f1", "F1", { recordType: "filter", league: null }),
    ];
    const allowed = new Set(records.map((record) => record.id));

    const results = searchRecordsByContext(records, allowed, "");

    expect(results.sourceOnlyResults).toHaveLength(0);
    expect(results.normalizedResults.map((result) => result.record.id)).toEqual([
      "entity:nba:e1",
      "filter:global:f1",
      "measure:mlb:m1",
      "measure:nba:m0",
    ]);
    expect(results.normalizedResults.every((result) => result.match_kind === "browse")).toBe(true);
  });

  it("builds deterministic source hints when normalized matches are zero", () => {
    const records: TypedRecord[] = [
      makeRecord("measure:mlb:strikeouts", "Strikeouts", {
        source_type: "lookup_index",
        source_ref: "lookup.playerSplits.bucket",
      }),
    ];
    const allowed = new Set(records.map((record) => record.id));
    const index = buildSemanticIndexFixture(records);

    const searchResults = searchRecordsByContext(records, allowed, "playerSplits");
    expect(searchResults.normalizedResults).toHaveLength(0);

    const narrative = buildSearchNarrative(
      index,
      "playerSplits",
      searchResults.normalizedResults,
      searchResults.sourceOnlyResults,
      allowed
    );

    expect(narrative).not.toBeNull();
    expect(narrative?.summary).toBe(
      "No normalized records matched this term. The term still appears in source-oriented data hints."
    );
    expect(narrative?.hints.map((hint) => `${hint.hint_type}:${hint.label}`)).toEqual([
      "record_source:measure:mlb:strikeouts",
      "source_path:lookup_index/mlb/playerSplits.lookup_index.json",
      "source_tree_label:playerSplits bucket",
    ]);
  });
});
