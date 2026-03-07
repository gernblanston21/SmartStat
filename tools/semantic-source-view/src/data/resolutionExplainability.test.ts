import { describe, expect, it } from "vitest";
import { buildResolutionExplainabilityModel } from "./resolutionExplainability";
import { RecordSearchResult } from "../types";

function makeResult(): RecordSearchResult {
  return {
    record: {
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      league: "mlb",
      source_type: "runtime",
      source_path: "runtime/leagues/mlb/fetchBaseMeasures.response.json",
      source_ref: "measures[].air_balls",
      aliases: ["air_b", "air_balls"],
      confidence: "high",
      notes: ["phase8 scaffold test"],
      lineage: [
        {
          source_type: "lookup_index",
          source_path: "lookup_index/mlb/lookup_index.json",
          source_ref: "measures.air_balls",
        },
        {
          source_type: "runtime",
          source_path: "runtime/leagues/mlb/fetchBaseMeasures.response.json",
          source_ref: "measures[].air_balls",
        },
      ],
      evidence: [
        "runtime/leagues/mlb/fetchBaseMeasures.response.json#measures[].air_balls",
        "lookup_index/mlb/lookup_index.json#measures.air_balls",
      ],
      related_ids: ["filter:mlb:last_10_games", "filter:mlb:season_to_date"],
      relationship_refs: ["rel:measure_to_filter:mlb:air_balls:last_10_games"],
      recordType: "measure",
    },
    match_kind: "exact_token",
    match_strength: "medium",
    match_fields: ["id", "name"],
    match_scope: "normalized",
    score: 760,
    explanation: "Matched exact normalized token",
  };
}

describe("resolution explainability scaffold", () => {
  it("builds deterministic models for identical inputs", () => {
    const input = makeResult();
    const modelA = buildResolutionExplainabilityModel(input, "air_balls", 2);
    const modelB = buildResolutionExplainabilityModel(input, "air_balls", 2);

    expect(modelA).toEqual(modelB);
  });

  it("uses fixed ordered step kinds", () => {
    const model = buildResolutionExplainabilityModel(makeResult(), "air_balls", 2);
    expect(model.steps.map((step) => step.kind)).toEqual([
      "search_entry",
      "semantic_scope",
      "source_lineage",
      "relationship_scope",
      "query_path_scope",
      "runtime_bridge_deferred",
    ]);
  });

  it("keeps lineage refs sorted and deduplicated", () => {
    const input = makeResult();
    input.record.lineage = [
      input.record.lineage[1],
      input.record.lineage[0],
      input.record.lineage[1],
    ];
    const model = buildResolutionExplainabilityModel(input, "air_balls", 1);
    const lineageStep = model.steps.find((step) => step.kind === "source_lineage");
    expect(lineageStep?.refs).toEqual([
      "lookup_index:lookup_index/mlb/lookup_index.json#measures.air_balls",
      "lookup_index/mlb/lookup_index.json#measures.air_balls",
      "runtime:runtime/leagues/mlb/fetchBaseMeasures.response.json#measures[].air_balls",
      "runtime/leagues/mlb/fetchBaseMeasures.response.json#measures[].air_balls",
    ]);
  });
});
