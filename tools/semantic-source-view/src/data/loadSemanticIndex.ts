import { SemanticIndex } from "../types";

const SEMANTIC_INDEX_ENDPOINT = "/api/semantic-index";

export async function loadSemanticIndex(): Promise<SemanticIndex> {
  const response = await fetch(SEMANTIC_INDEX_ENDPOINT, {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch semantic index (${response.status})`);
  }

  return (await response.json()) as SemanticIndex;
}
