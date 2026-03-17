import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  PreviewPayloadContractError,
  adaptPreviewPayloadToViewModel,
} from "../adapters/previewPayloadToViewModel";

type JsonMap = Record<string, unknown>;

function findRepoRoot(startDir: string): string {
  let current = resolve(startDir);
  while (true) {
    if (existsSync(join(current, "AGENTS.md"))) {
      return current;
    }
    const parent = dirname(current);
    if (parent === current) {
      throw new Error(`Unable to locate repo root from ${startDir}`);
    }
    current = parent;
  }
}

function readJsonFile(path: string): JsonMap {
  return JSON.parse(readFileSync(path, "utf8")) as JsonMap;
}

function deepClone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function sha256(text: string): string {
  return createHash("sha256").update(text).digest("hex");
}

const THIS_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = findRepoRoot(THIS_DIR);
const RUN1_PATH = join(
  REPO_ROOT,
  "tests",
  "_scratch",
  "runtime-slice-02-readonly-plan-bridge",
  "runs",
  "pos_projection_intake_run1.json"
);
const RUN2_PATH = join(
  REPO_ROOT,
  "tests",
  "_scratch",
  "runtime-slice-02-readonly-plan-bridge",
  "runs",
  "pos_projection_intake_run2.json"
);

describe("previewPayloadToViewModel adapter", () => {
  it("produces byte-stable deterministic output for run1 and run2", () => {
    const run1 = readJsonFile(RUN1_PATH);
    const run2 = readJsonFile(RUN2_PATH);

    const adapted1 = adaptPreviewPayloadToViewModel(run1);
    const adapted2 = adaptPreviewPayloadToViewModel(run2);

    const json1 = JSON.stringify(adapted1);
    const json2 = JSON.stringify(adapted2);

    expect(adapted1).toEqual(adapted2);
    expect(json1).toBe(json2);
    expect(sha256(json1)).toBe(sha256(json2));

    expect(Object.keys(adapted1)).toEqual([
      "adapter_contract",
      "source_preview_kind",
      "source_payload_kind",
      "view_model",
    ]);
    expect(Object.keys(adapted1.view_model)).toEqual([
      "intake_header",
      "projection_metadata_view",
      "semantic_view",
      "issues_view",
      "resolution_view",
      "rule_evaluation_summary_view",
      "rule_evaluation_trace_view",
      "deterministic_identity_view",
      "raw_payload_debug_view",
    ]);
  });

  it("fails closed for malformed payload values", () => {
    const malformed = deepClone(readJsonFile(RUN1_PATH));
    const payload = malformed.preview_payload as JsonMap;
    const trace = payload.rule_evaluation_trace_preview as JsonMap;
    const deterministicIdentity = trace.deterministic_identity_summary as JsonMap;
    deterministicIdentity.replay_identity = "";

    try {
      adaptPreviewPayloadToViewModel(malformed);
      throw new Error("expected adapter to fail closed");
    } catch (error: unknown) {
      expect(error).toBeInstanceOf(PreviewPayloadContractError);
      const contractError = error as PreviewPayloadContractError;
      expect(contractError.code).toBe("VIEWER_PREVIEW_CONTRACT_VIOLATION");
      expect(contractError.path).toContain(
        "rule_evaluation_trace_preview.deterministic_identity_summary.replay_identity"
      );
    }
  });

  it("fails closed when required fields are missing", () => {
    const missing = deepClone(readJsonFile(RUN1_PATH));
    const payload = missing.preview_payload as JsonMap;
    const projection = payload.projection_metadata as JsonMap;
    delete projection.status_summary;

    expect(() => adaptPreviewPayloadToViewModel(missing)).toThrowError(
      PreviewPayloadContractError
    );
  });

  it("does not leak unapproved input fields into output", () => {
    const withExtras = deepClone(readJsonFile(RUN1_PATH));
    (withExtras as JsonMap).runtime_internal_debug = {
      should_not_surface: true,
    };
    const payload = withExtras.preview_payload as JsonMap;
    payload.unapproved_runtime_section = { unsafe: true };

    const adapted = adaptPreviewPayloadToViewModel(withExtras);
    const outputJson = JSON.stringify(adapted);

    expect(outputJson).not.toContain("runtime_internal_debug");
    expect(outputJson).not.toContain("unapproved_runtime_section");
  });
});
