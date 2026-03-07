import { RecordSearchResult, SearchMatchKind, SearchMatchStrength } from "../types";

export const SEARCH_MATCH_SCORE: Record<SearchMatchKind, number> = {
  browse: 0,
  exact_id: 1000,
  exact_name: 990,
  exact_alias: 980,
  prefix: 860,
  exact_token: 760,
  substring: 640,
  source_ref: 420,
  source_path: 410,
  evidence: 400,
};

export const SOURCE_ONLY_MATCH_KINDS = new Set<SearchMatchKind>([
  "source_ref",
  "source_path",
  "evidence",
]);

export function normalizeSearchText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_");
}

export function splitSearchTokens(value: string): string[] {
  return value.split(/[^a-z0-9]+/g).filter(Boolean);
}

export function compareRecordSearchResults(left: RecordSearchResult, right: RecordSearchResult): number {
  if (left.score !== right.score) {
    return right.score - left.score;
  }

  const typeCmp = left.record.recordType.localeCompare(right.record.recordType);
  if (typeCmp !== 0) {
    return typeCmp;
  }

  const leftLeague = left.record.league ?? "~";
  const rightLeague = right.record.league ?? "~";
  const leagueCmp = leftLeague.localeCompare(rightLeague);
  if (leagueCmp !== 0) {
    return leagueCmp;
  }

  return left.record.id.localeCompare(right.record.id);
}

export function matchStrengthForKind(kind: SearchMatchKind): SearchMatchStrength {
  if (kind === "browse") {
    return "none";
  }
  if (kind === "exact_id" || kind === "exact_name" || kind === "exact_alias") {
    return "high";
  }
  if (kind === "prefix" || kind === "exact_token" || kind === "substring") {
    return "medium";
  }
  return "low";
}

export function matchLabel(kind: SearchMatchKind): string {
  if (kind === "exact_id" || kind === "exact_name" || kind === "exact_alias") {
    return "exact";
  }
  if (kind === "prefix") {
    return "prefix";
  }
  if (kind === "exact_token") {
    return "token";
  }
  if (kind === "substring") {
    return "substring";
  }
  if (kind === "browse") {
    return "browse";
  }
  return "source";
}
