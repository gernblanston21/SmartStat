import { RecordSearchResult } from "../types";
import { compareRecordSearchResults, normalizeSearchText } from "./searchContract";

export const CANDIDATE_RESOLUTION_SCHEMA_VERSION = "candidate-resolution.v0";

export type CandidateResolutionStatus = "preferred" | "alternate" | "rejected";

export interface CandidateResolutionEntry {
  candidate_id: string;
  record_type: string;
  league: string | null;
  ranking_position: number | null;
  status: CandidateResolutionStatus;
  match_kind: string;
  match_strength: string;
  reason: string;
  source_type: string;
  source_path: string;
  source_ref: string;
  supporting_refs: string[];
}

export interface CandidateResolutionModel {
  schema_version: string;
  mode: "read_only_scaffold";
  input_term: string;
  normalized_input_term: string;
  preferred_candidate_id: string | null;
  ambiguity_flag: boolean;
  ambiguity_reason: string | null;
  ranking_contract: string;
  ranking_order: string[];
  candidates: CandidateResolutionEntry[];
  deferred_boundaries: string[];
}

export interface CandidateResolutionViewModel {
  model: CandidateResolutionModel;
  model_source: "real_data" | "fixture";
}

const DEFAULT_REJECTED_LIMIT = 5;

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values)).sort((left, right) => left.localeCompare(right));
}

function toSourceRef(result: RecordSearchResult): string {
  return `${result.record.source_type}:${result.record.source_path}#${result.record.source_ref}`;
}

function buildSupportingRefs(result: RecordSearchResult): string[] {
  const lineageRefs = result.record.lineage.map(
    (entry) => `${entry.source_type}:${entry.source_path}#${entry.source_ref}`
  );
  return uniqueSorted([toSourceRef(result), ...lineageRefs, ...result.record.evidence]);
}

function toEntry(
  result: RecordSearchResult,
  status: CandidateResolutionStatus,
  rankingPosition: number | null,
  reason: string
): CandidateResolutionEntry {
  return {
    candidate_id: result.record.id,
    record_type: result.record.recordType,
    league: result.record.league,
    ranking_position: rankingPosition,
    status,
    match_kind: result.match_kind,
    match_strength: result.match_strength,
    reason,
    source_type: result.record.source_type,
    source_path: result.record.source_path,
    source_ref: result.record.source_ref,
    supporting_refs: buildSupportingRefs(result),
  };
}

function buildPreferredReason(isTopRanked: boolean): string {
  if (isTopRanked) {
    return "Top deterministic normalized candidate for this search context.";
  }
  return "Developer-selected candidate pinned for deterministic inspection context.";
}

export function buildCandidateResolutionModel(params: {
  searchTerm: string;
  selectedResult: RecordSearchResult;
  normalizedResults: RecordSearchResult[];
  sourceOnlyResults: RecordSearchResult[];
  rejectedLimit?: number;
}): CandidateResolutionModel {
  const trimmedTerm = params.searchTerm.trim();
  const normalizedTerm = normalizeSearchText(trimmedTerm);
  const rejectedLimit = params.rejectedLimit ?? DEFAULT_REJECTED_LIMIT;

  const rankedNormalized = params.normalizedResults.slice().sort(compareRecordSearchResults);
  const selectedInRanked = rankedNormalized.find(
    (result) => result.record.id === params.selectedResult.record.id
  );
  const preferredResult = selectedInRanked ?? params.selectedResult;

  const topRankedId = rankedNormalized[0]?.record.id ?? null;
  const candidates: CandidateResolutionEntry[] = [];
  candidates.push(
    toEntry(
      preferredResult,
      "preferred",
      1,
      buildPreferredReason(topRankedId === preferredResult.record.id)
    )
  );

  const alternateResults = rankedNormalized.filter(
    (result) => result.record.id !== preferredResult.record.id
  );
  for (const [index, result] of alternateResults.entries()) {
    candidates.push(
      toEntry(
        result,
        "alternate",
        index + 2,
        "Normalized candidate with lower deterministic rank than preferred candidate."
      )
    );
  }

  const rankedSourceOnly = params.sourceOnlyResults.slice().sort(compareRecordSearchResults);
  const rejectedResults = rankedSourceOnly
    .filter((result) => !candidates.some((candidate) => candidate.candidate_id === result.record.id))
    .slice(0, rejectedLimit);
  for (const result of rejectedResults) {
    candidates.push(
      toEntry(
        result,
        "rejected",
        null,
        "Source-only match; excluded from normalized candidate set in this Phase 9 scaffold."
      )
    );
  }

  const equalScoreCandidates = rankedNormalized.filter(
    (result) => result.score === preferredResult.score
  );
  const ambiguityFlag = equalScoreCandidates.length > 1;
  const ambiguityReason = ambiguityFlag
    ? `Multiple normalized candidates share deterministic score ${preferredResult.score}.`
    : null;

  return {
    schema_version: CANDIDATE_RESOLUTION_SCHEMA_VERSION,
    mode: "read_only_scaffold",
    input_term: trimmedTerm,
    normalized_input_term: normalizedTerm,
    preferred_candidate_id: preferredResult.record.id,
    ambiguity_flag: ambiguityFlag,
    ambiguity_reason: ambiguityReason,
    ranking_contract: "phase9_semantic_candidate_scaffold_v1",
    ranking_order: [
      "match_score(desc)",
      "record_type(asc)",
      "league(asc)",
      "record_id(asc)",
      "selected_record_preferred_override",
    ],
    candidates,
    deferred_boundaries: [
      "runtime apply behavior",
      "runtime resolver integration",
      "planner execution",
      "plan generation",
      "execution bridge behavior",
    ],
  };
}

export function resolveCandidateResolutionViewModel(params: {
  searchTerm: string;
  selectedResult: RecordSearchResult | null;
  normalizedResults: RecordSearchResult[];
  sourceOnlyResults: RecordSearchResult[];
  fixtureModel: CandidateResolutionModel;
}): CandidateResolutionViewModel {
  if (!params.selectedResult) {
    return {
      model: params.fixtureModel,
      model_source: "fixture",
    };
  }

  return {
    model: buildCandidateResolutionModel({
      searchTerm: params.searchTerm,
      selectedResult: params.selectedResult,
      normalizedResults: params.normalizedResults,
      sourceOnlyResults: params.sourceOnlyResults,
    }),
    model_source: "real_data",
  };
}
