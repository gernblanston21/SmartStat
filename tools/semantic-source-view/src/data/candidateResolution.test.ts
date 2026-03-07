import { describe, expect, it } from "vitest";
import {
  buildCandidateResolutionModel,
  resolveCandidateResolutionViewModel,
} from "./candidateResolution";
import { PHASE9_CANDIDATE_RESOLUTION_FIXTURE } from "./candidateResolution.fixture";
import { RecordSearchResult } from "../types";

function makeResult(params: {
  id: string;
  name: string;
  recordType?: "entity" | "measure" | "filter" | "profile" | "qualifier";
  league?: string | null;
  matchKind?: RecordSearchResult["match_kind"];
  score?: number;
  sourceType?: RecordSearchResult["record"]["source_type"];
  sourceRef?: string;
}): RecordSearchResult {
  const recordType = params.recordType ?? "measure";
  const sourceType = params.sourceType ?? "runtime";
  return {
    record: {
      id: params.id,
      name: params.name,
      league: params.league ?? "mlb",
      source_type: sourceType,
      source_path:
        sourceType === "runtime"
          ? "runtime/leagues/mlb/fetchBaseMeasures.response.json"
          : "lookup_index/mlb/lookup_index.json",
      source_ref: params.sourceRef ?? "measures[]",
      aliases: [],
      confidence: "high",
      notes: [],
      lineage: [
        {
          source_type: sourceType,
          source_path:
            sourceType === "runtime"
              ? "runtime/leagues/mlb/fetchBaseMeasures.response.json"
              : "lookup_index/mlb/lookup_index.json",
          source_ref: params.sourceRef ?? "measures[]",
        },
      ],
      evidence: [],
      related_ids: [],
      relationship_refs: [],
      recordType,
    },
    match_kind: params.matchKind ?? "substring",
    match_strength: "medium",
    match_fields: ["id"],
    match_scope: sourceType === "runtime" ? "normalized" : "source_hint",
    score: params.score ?? 640,
    explanation: "test match",
  };
}

describe("phase9 candidate-resolution scaffold", () => {
  it("builds identical models for identical input", () => {
    const selected = makeResult({
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      matchKind: "exact_token",
      score: 760,
    });
    const alternates = [
      selected,
      makeResult({
        id: "measure:mlb:air_balls_percentage",
        name: "Air Balls Percentage",
        matchKind: "substring",
        score: 640,
      }),
    ];
    const sourceOnly = [
      makeResult({
        id: "measure:mlb:air_balls_lookup",
        name: "Air Balls Lookup",
        sourceType: "lookup_index",
        matchKind: "source_ref",
        score: 420,
      }),
    ];

    const modelA = buildCandidateResolutionModel({
      searchTerm: "air_balls",
      selectedResult: selected,
      normalizedResults: alternates,
      sourceOnlyResults: sourceOnly,
    });
    const modelB = buildCandidateResolutionModel({
      searchTerm: "air_balls",
      selectedResult: selected,
      normalizedResults: alternates,
      sourceOnlyResults: sourceOnly,
    });

    expect(modelA).toEqual(modelB);
  });

  it("keeps deterministic candidate ordering", () => {
    const preferred = makeResult({
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      matchKind: "exact_token",
      score: 760,
    });
    const alternate = makeResult({
      id: "measure:mlb:air_balls_percentage",
      name: "Air Balls Percentage",
      matchKind: "substring",
      score: 640,
    });
    const rejected = makeResult({
      id: "measure:mlb:air_balls_lookup",
      name: "Air Balls Lookup",
      sourceType: "lookup_index",
      matchKind: "source_ref",
      score: 420,
    });

    const model = buildCandidateResolutionModel({
      searchTerm: "air_balls",
      selectedResult: preferred,
      normalizedResults: [alternate, preferred],
      sourceOnlyResults: [rejected],
    });

    expect(model.candidates.map((candidate) => candidate.candidate_id)).toEqual([
      "measure:mlb:air_balls",
      "measure:mlb:air_balls_percentage",
      "measure:mlb:air_balls_lookup",
    ]);
  });

  it("applies stable preferred/alternate/rejected statuses", () => {
    const selected = makeResult({
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      matchKind: "exact_token",
      score: 760,
    });
    const alternate = makeResult({
      id: "measure:mlb:air_balls_percentage",
      name: "Air Balls Percentage",
      matchKind: "substring",
      score: 640,
    });
    const rejected = makeResult({
      id: "measure:mlb:air_balls_lookup",
      name: "Air Balls Lookup",
      sourceType: "lookup_index",
      matchKind: "source_path",
      score: 410,
    });

    const model = buildCandidateResolutionModel({
      searchTerm: "air_balls",
      selectedResult: selected,
      normalizedResults: [selected, alternate],
      sourceOnlyResults: [rejected],
    });

    expect(
      model.candidates.map((candidate) => ({
        id: candidate.candidate_id,
        status: candidate.status,
        rank: candidate.ranking_position,
      }))
    ).toEqual([
      { id: "measure:mlb:air_balls", status: "preferred", rank: 1 },
      { id: "measure:mlb:air_balls_percentage", status: "alternate", rank: 2 },
      { id: "measure:mlb:air_balls_lookup", status: "rejected", rank: null },
    ]);
  });

  it("sets ambiguity deterministically when top scores tie", () => {
    const selected = makeResult({
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      matchKind: "exact_token",
      score: 760,
    });
    const tie = makeResult({
      id: "measure:mlb:air_balls_rate",
      name: "Air Balls Rate",
      matchKind: "exact_token",
      score: 760,
    });

    const model = buildCandidateResolutionModel({
      searchTerm: "air_balls",
      selectedResult: selected,
      normalizedResults: [selected, tie],
      sourceOnlyResults: [],
    });

    expect(model.ambiguity_flag).toBe(true);
    expect(model.ambiguity_reason).toBe("Multiple normalized candidates share deterministic score 760.");
  });

  it("uses real-data model path when selected data is available", () => {
    const selected = makeResult({
      id: "measure:mlb:air_balls",
      name: "Air Balls",
      matchKind: "exact_token",
      score: 760,
    });
    const alternate = makeResult({
      id: "measure:mlb:air_balls_percentage",
      name: "Air Balls Percentage",
      matchKind: "substring",
      score: 640,
    });

    const viewModel = resolveCandidateResolutionViewModel({
      searchTerm: "air_balls",
      selectedResult: selected,
      normalizedResults: [selected, alternate],
      sourceOnlyResults: [],
      fixtureModel: PHASE9_CANDIDATE_RESOLUTION_FIXTURE,
    });

    expect(viewModel.model_source).toBe("real_data");
    expect(viewModel.model.preferred_candidate_id).toBe("measure:mlb:air_balls");
    expect(viewModel.model).not.toEqual(PHASE9_CANDIDATE_RESOLUTION_FIXTURE);
  });
});
