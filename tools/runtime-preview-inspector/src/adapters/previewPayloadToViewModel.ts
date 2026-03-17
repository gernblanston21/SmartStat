import {
  EXPECTED_BRIDGE_MODE,
  EXPECTED_PAYLOAD_KIND,
  EXPECTED_PREVIEW_KIND,
  EXPECTED_RULE_PHASE_ORDER,
  FieldPreviewEntry,
  ProjectionMetadata,
  REQUIRED_TRACEABILITY_OVERLAP_FIELDS,
  ResolutionPreview,
  RuleEvaluationSummaryPreview,
  RuleEvaluationTracePreview,
  RulePhase,
  RuntimePreviewViewModel,
  RUNTIME_PREVIEW_VIEW_MODEL_CONTRACT,
  SemanticInterpretationSummary,
} from "../contracts/runtimePreviewIntake";

type JsonObject = Record<string, unknown>;

const SHA256_HEX_64 = /^[a-fA-F0-9]{64}$/;

export class PreviewPayloadContractError extends Error {
  readonly code = "VIEWER_PREVIEW_CONTRACT_VIOLATION";
  readonly path: string;

  constructor(path: string, detail: string) {
    super(`VIEWER_PREVIEW_CONTRACT_VIOLATION at ${path}: ${detail}`);
    this.name = "PreviewPayloadContractError";
    this.path = path;
  }
}

function fail(path: string, detail: string): never {
  throw new PreviewPayloadContractError(path, detail);
}

function asObject(value: unknown, path: string): JsonObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail(path, "expected object");
  }
  return value as JsonObject;
}

function readRequired(value: JsonObject, key: string, path: string): unknown {
  if (!Object.prototype.hasOwnProperty.call(value, key)) {
    fail(`${path}.${key}`, "missing required key");
  }
  return value[key];
}

function readString(
  value: JsonObject,
  key: string,
  path: string,
  allowEmpty = false
): string {
  const raw = readRequired(value, key, path);
  if (typeof raw !== "string") {
    fail(`${path}.${key}`, "expected string");
  }
  if (!allowEmpty && raw.trim().length === 0) {
    fail(`${path}.${key}`, "expected non-empty string");
  }
  return raw;
}

function readBoolean(value: JsonObject, key: string, path: string): boolean {
  const raw = readRequired(value, key, path);
  if (typeof raw !== "boolean") {
    fail(`${path}.${key}`, "expected boolean");
  }
  return raw;
}

function readNonNegativeInteger(
  value: JsonObject,
  key: string,
  path: string
): number {
  const raw = readRequired(value, key, path);
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    fail(`${path}.${key}`, "expected non-negative integer");
  }
  return raw;
}

function readArray(value: JsonObject, key: string, path: string): unknown[] {
  const raw = readRequired(value, key, path);
  if (!Array.isArray(raw)) {
    fail(`${path}.${key}`, "expected array");
  }
  return raw;
}

function readStringArray(
  value: JsonObject,
  key: string,
  path: string,
  allowEmptyValues = false
): string[] {
  const raw = readArray(value, key, path);
  return raw.map((entry, index) => {
    if (typeof entry !== "string") {
      fail(`${path}.${key}[${index}]`, "expected string");
    }
    if (!allowEmptyValues && entry.trim().length === 0) {
      fail(`${path}.${key}[${index}]`, "expected non-empty string");
    }
    return entry;
  });
}

function readSha256Hex(value: JsonObject, key: string, path: string): string {
  const parsed = readString(value, key, path);
  if (!SHA256_HEX_64.test(parsed)) {
    fail(`${path}.${key}`, "expected 64-char sha256 hex");
  }
  return parsed.toLowerCase();
}

function assertRelativeKeyOrder(
  value: JsonObject,
  requiredKeys: readonly string[],
  path: string
): void {
  const actual = Object.keys(value).filter((key) => requiredKeys.includes(key));
  if (actual.length !== requiredKeys.length) {
    fail(path, "required keys missing for order check");
  }
  for (let index = 0; index < requiredKeys.length; index += 1) {
    if (actual[index] !== requiredKeys[index]) {
      fail(
        path,
        `required key order mismatch: expected ${requiredKeys.join(
          ", "
        )}; got ${actual.join(", ")}`
      );
    }
  }
}

function deepCloneArray(input: unknown[]): unknown[] {
  return JSON.parse(JSON.stringify(input)) as unknown[];
}

function parseInputIdentity(value: JsonObject, path: string): {
  artifact_path: string;
  input_fingerprint_sha256: string;
} {
  return {
    artifact_path: readString(value, "artifact_path", path),
    input_fingerprint_sha256: readSha256Hex(
      value,
      "input_fingerprint_sha256",
      path
    ),
  };
}

function parseDeterministicIdentitySummary(
  value: JsonObject,
  path: string
): {
  normalized_plan_hash: string;
  replay_identity: string;
  validator_run_identity: string;
} {
  return {
    normalized_plan_hash: readSha256Hex(value, "normalized_plan_hash", path),
    replay_identity: readSha256Hex(value, "replay_identity", path),
    validator_run_identity: readSha256Hex(value, "validator_run_identity", path),
  };
}

function parseSemanticInterpretationSummary(
  value: JsonObject,
  path: string
): SemanticInterpretationSummary {
  return {
    scope_resolution: readString(value, "scope_resolution", path),
    effective_scope: readString(value, "effective_scope", path),
    evidence_source: readString(value, "evidence_source", path),
  };
}

function parseProjectionMetadata(
  value: JsonObject,
  path: string
): ProjectionMetadata {
  assertRelativeKeyOrder(
    value,
    [
      "projection_contract",
      "projection_kind",
      "input_artifact",
      "input_identity",
      "status_summary",
      "deterministic_identity_summary",
      "semantic_interpretation_summary",
      "issues_summary",
    ],
    path
  );

  const inputIdentity = parseInputIdentity(
    asObject(readRequired(value, "input_identity", path), `${path}.input_identity`),
    `${path}.input_identity`
  );

  const statusSummaryObject = asObject(
    readRequired(value, "status_summary", path),
    `${path}.status_summary`
  );
  const status = readString(statusSummaryObject, "status", `${path}.status_summary`);
  if (status !== "PASS" && status !== "REFUSE") {
    fail(`${path}.status_summary.status`, "expected PASS or REFUSE");
  }

  const deterministicIdentitySummary = parseDeterministicIdentitySummary(
    asObject(
      readRequired(value, "deterministic_identity_summary", path),
      `${path}.deterministic_identity_summary`
    ),
    `${path}.deterministic_identity_summary`
  );

  const semanticInterpretationSummary = parseSemanticInterpretationSummary(
    asObject(
      readRequired(value, "semantic_interpretation_summary", path),
      `${path}.semantic_interpretation_summary`
    ),
    `${path}.semantic_interpretation_summary`
  );

  const issuesSummaryObject = asObject(
    readRequired(value, "issues_summary", path),
    `${path}.issues_summary`
  );
  const errors = deepCloneArray(readArray(issuesSummaryObject, "errors", `${path}.issues_summary`));
  const warnings = deepCloneArray(
    readArray(issuesSummaryObject, "warnings", `${path}.issues_summary`)
  );

  return {
    projection_contract: readString(value, "projection_contract", path),
    projection_kind: readString(value, "projection_kind", path),
    input_artifact: readString(value, "input_artifact", path),
    input_identity: inputIdentity,
    status_summary: {
      status,
      error_count: readNonNegativeInteger(
        statusSummaryObject,
        "error_count",
        `${path}.status_summary`
      ),
      warning_count: readNonNegativeInteger(
        statusSummaryObject,
        "warning_count",
        `${path}.status_summary`
      ),
    },
    deterministic_identity_summary: deterministicIdentitySummary,
    semantic_interpretation_summary: semanticInterpretationSummary,
    issues_summary: {
      errors,
      warnings,
    },
  };
}

function parseResolutionPreview(value: JsonObject, path: string): ResolutionPreview {
  const status = readString(value, "status", path);
  if (status !== "PASS" && status !== "REFUSE") {
    fail(`${path}.status`, "expected PASS or REFUSE");
  }
  return {
    status,
    scope_resolution: readString(value, "scope_resolution", path),
    effective_scope: readString(value, "effective_scope", path),
    evidence_source: readString(value, "evidence_source", path),
  };
}

function parseRuleEvaluationSummaryPreview(
  value: JsonObject,
  path: string
): RuleEvaluationSummaryPreview {
  const phaseOrder = readStringArray(value, "phase_order", path) as RulePhase[];
  if (phaseOrder.length !== EXPECTED_RULE_PHASE_ORDER.length) {
    fail(
      `${path}.phase_order`,
      `expected ${EXPECTED_RULE_PHASE_ORDER.length} phases`
    );
  }
  for (let index = 0; index < EXPECTED_RULE_PHASE_ORDER.length; index += 1) {
    if (phaseOrder[index] !== EXPECTED_RULE_PHASE_ORDER[index]) {
      fail(
        `${path}.phase_order[${index}]`,
        `expected ${EXPECTED_RULE_PHASE_ORDER[index]}`
      );
    }
  }

  const orderedRulesRaw = readArray(value, "ordered_rules", path);
  const orderedRules = orderedRulesRaw.map((entry, index) => {
    const rulePath = `${path}.ordered_rules[${index}]`;
    const ruleObject = asObject(entry, rulePath);
    const category = readString(ruleObject, "category", rulePath);
    if (!phaseOrder.includes(category as RulePhase)) {
      fail(`${rulePath}.category`, `unknown phase category: ${category}`);
    }
    return {
      category: category as RulePhase,
      rule_id: readString(ruleObject, "rule_id", rulePath),
      outcome: readString(ruleObject, "outcome", rulePath),
    };
  });

  let previousIndex = 0;
  orderedRules.forEach((rule, index) => {
    const currentIndex = phaseOrder.indexOf(rule.category);
    if (index > 0 && currentIndex < previousIndex) {
      fail(
        `${path}.ordered_rules[${index}].category`,
        "rule category ordering regressed versus phase_order"
      );
    }
    previousIndex = currentIndex;
  });

  return {
    phase_order: [...phaseOrder],
    ordered_rules: orderedRules,
  };
}

function parseRuleEvaluationTracePreview(
  value: JsonObject,
  path: string
): RuleEvaluationTracePreview {
  return {
    projection_contract: readString(value, "projection_contract", path),
    projection_kind: readString(value, "projection_kind", path),
    input_artifact: readString(value, "input_artifact", path),
    input_identity: parseInputIdentity(
      asObject(readRequired(value, "input_identity", path), `${path}.input_identity`),
      `${path}.input_identity`
    ),
    deterministic_identity_summary: parseDeterministicIdentitySummary(
      asObject(
        readRequired(value, "deterministic_identity_summary", path),
        `${path}.deterministic_identity_summary`
      ),
      `${path}.deterministic_identity_summary`
    ),
  };
}

function parseFieldPreview(value: JsonObject, path: string): FieldPreviewEntry[] {
  const raw = readArray(value, "field_preview", path);
  return raw.map((entry, index) => {
    const itemPath = `${path}.field_preview[${index}]`;
    const objectValue = asObject(entry, itemPath);
    return {
      name: readString(objectValue, "name", itemPath),
      page_property: readString(objectValue, "page_property", itemPath, true),
      custom_property: readString(objectValue, "custom_property", itemPath, true),
    };
  });
}

function normalizeFieldPreview(
  fieldPreview: FieldPreviewEntry[],
  tabfieldOrder: string[],
  path: string
): FieldPreviewEntry[] {
  const byName = new Map<string, FieldPreviewEntry>();
  fieldPreview.forEach((entry, index) => {
    if (byName.has(entry.name)) {
      fail(`${path}[${index}].name`, "duplicate field_preview.name");
    }
    byName.set(entry.name, {
      name: entry.name,
      page_property: entry.page_property,
      custom_property: entry.custom_property,
    });
  });

  const normalized = tabfieldOrder.map((tabfieldName, index) => {
    const found = byName.get(tabfieldName);
    if (!found) {
      fail(`${path}`, `missing field_preview entry for tabfield_order[${index}]`);
    }
    return found;
  });

  if (normalized.length !== byName.size) {
    fail(path, "field_preview contains names not present in tabfield_order");
  }
  return normalized;
}

function assertProjectionTraceOverlap(
  projectionMetadata: ProjectionMetadata,
  tracePreview: RuleEvaluationTracePreview,
  path: string
): void {
  const overlapMismatch = REQUIRED_TRACEABILITY_OVERLAP_FIELDS.find((fieldName) => {
    if (fieldName === "input_identity") {
      return (
        projectionMetadata.input_identity.artifact_path !==
          tracePreview.input_identity.artifact_path ||
        projectionMetadata.input_identity.input_fingerprint_sha256 !==
          tracePreview.input_identity.input_fingerprint_sha256
      );
    }
    if (fieldName === "deterministic_identity_summary") {
      return (
        projectionMetadata.deterministic_identity_summary.normalized_plan_hash !==
          tracePreview.deterministic_identity_summary.normalized_plan_hash ||
        projectionMetadata.deterministic_identity_summary.replay_identity !==
          tracePreview.deterministic_identity_summary.replay_identity ||
        projectionMetadata.deterministic_identity_summary.validator_run_identity !==
          tracePreview.deterministic_identity_summary.validator_run_identity
      );
    }
    return projectionMetadata[fieldName] !== tracePreview[fieldName];
  });

  if (overlapMismatch) {
    fail(
      path,
      `projection_metadata and rule_evaluation_trace_preview mismatch at ${overlapMismatch}`
    );
  }
}

function assertResolutionConsistency(
  projectionMetadata: ProjectionMetadata,
  resolutionPreview: ResolutionPreview,
  path: string
): void {
  if (projectionMetadata.status_summary.status !== resolutionPreview.status) {
    fail(path, "resolution_preview.status mismatch with status_summary.status");
  }
  if (
    projectionMetadata.semantic_interpretation_summary.scope_resolution !==
      resolutionPreview.scope_resolution ||
    projectionMetadata.semantic_interpretation_summary.effective_scope !==
      resolutionPreview.effective_scope ||
    projectionMetadata.semantic_interpretation_summary.evidence_source !==
      resolutionPreview.evidence_source
  ) {
    fail(
      path,
      "resolution_preview semantic fields mismatch semantic_interpretation_summary"
    );
  }
}

export function adaptPreviewPayloadToViewModel(
  artifactJson: unknown
): RuntimePreviewViewModel {
  const root = asObject(artifactJson, "$");

  const status = readString(root, "status", "$");
  if (status !== "success") {
    fail("$.status", "adapter requires success status artifact");
  }
  const previewKind = readString(root, "preview_kind", "$");
  if (previewKind !== EXPECTED_PREVIEW_KIND) {
    fail(
      "$.preview_kind",
      `unsupported preview_kind; expected ${EXPECTED_PREVIEW_KIND}`
    );
  }

  const previewPayload = asObject(readRequired(root, "preview_payload", "$"), "$.preview_payload");
  const payloadKind = readString(previewPayload, "payload_kind", "$.preview_payload");
  if (payloadKind !== EXPECTED_PAYLOAD_KIND) {
    fail(
      "$.preview_payload.payload_kind",
      `unsupported payload_kind; expected ${EXPECTED_PAYLOAD_KIND}`
    );
  }

  const bridgeMode = readString(previewPayload, "bridge_mode", "$.preview_payload");
  if (bridgeMode !== EXPECTED_BRIDGE_MODE) {
    fail(
      "$.preview_payload.bridge_mode",
      `unsupported bridge_mode; expected ${EXPECTED_BRIDGE_MODE}`
    );
  }

  const mutationAuthorized = readBoolean(
    previewPayload,
    "mutation_authorized",
    "$.preview_payload"
  );
  if (mutationAuthorized) {
    fail("$.preview_payload.mutation_authorized", "must be false for read-only viewer");
  }

  const tabfieldCount = readNonNegativeInteger(
    previewPayload,
    "tabfield_count",
    "$.preview_payload"
  );
  const tabfieldOrder = readStringArray(
    previewPayload,
    "tabfield_order",
    "$.preview_payload"
  );
  const fieldPreview = parseFieldPreview(previewPayload, "$.preview_payload");
  const normalizedFieldPreview = normalizeFieldPreview(
    fieldPreview,
    tabfieldOrder,
    "$.preview_payload.field_preview"
  );

  if (
    tabfieldCount !== tabfieldOrder.length ||
    tabfieldCount !== normalizedFieldPreview.length
  ) {
    fail(
      "$.preview_payload.tabfield_count",
      "tabfield_count must match tabfield_order and normalized field_preview lengths"
    );
  }

  const projectionMetadata = parseProjectionMetadata(
    asObject(
      readRequired(previewPayload, "projection_metadata", "$.preview_payload"),
      "$.preview_payload.projection_metadata"
    ),
    "$.preview_payload.projection_metadata"
  );
  const resolutionPreview = parseResolutionPreview(
    asObject(
      readRequired(previewPayload, "resolution_preview", "$.preview_payload"),
      "$.preview_payload.resolution_preview"
    ),
    "$.preview_payload.resolution_preview"
  );
  const ruleEvaluationSummaryPreview = parseRuleEvaluationSummaryPreview(
    asObject(
      readRequired(
        previewPayload,
        "rule_evaluation_summary_preview",
        "$.preview_payload"
      ),
      "$.preview_payload.rule_evaluation_summary_preview"
    ),
    "$.preview_payload.rule_evaluation_summary_preview"
  );
  const ruleEvaluationTracePreview = parseRuleEvaluationTracePreview(
    asObject(
      readRequired(
        previewPayload,
        "rule_evaluation_trace_preview",
        "$.preview_payload"
      ),
      "$.preview_payload.rule_evaluation_trace_preview"
    ),
    "$.preview_payload.rule_evaluation_trace_preview"
  );

  assertProjectionTraceOverlap(
    projectionMetadata,
    ruleEvaluationTracePreview,
    "$.preview_payload"
  );
  assertResolutionConsistency(
    projectionMetadata,
    resolutionPreview,
    "$.preview_payload.resolution_preview"
  );

  return {
    adapter_contract: RUNTIME_PREVIEW_VIEW_MODEL_CONTRACT,
    source_preview_kind: previewKind,
    source_payload_kind: payloadKind,
    view_model: {
      intake_header: {
        status,
        slice_name: readString(root, "slice_name", "$"),
        preview_kind: previewKind,
        provider_mode: readString(root, "provider_mode", "$"),
        normalization_rule: readString(root, "normalization_rule", "$"),
        page_name: readString(root, "page_name", "$"),
        page_template: readString(root, "page_template", "$"),
        supported_read_surfaces: readStringArray(
          root,
          "supported_read_surfaces",
          "$"
        ),
        bridge_mode: bridgeMode,
        mutation_authorized: mutationAuthorized,
      },
      projection_metadata_view: {
        projection_contract: projectionMetadata.projection_contract,
        projection_kind: projectionMetadata.projection_kind,
        input_artifact: projectionMetadata.input_artifact,
        input_identity: {
          artifact_path: projectionMetadata.input_identity.artifact_path,
          input_fingerprint_sha256:
            projectionMetadata.input_identity.input_fingerprint_sha256,
        },
        status_summary: {
          status: projectionMetadata.status_summary.status,
          error_count: projectionMetadata.status_summary.error_count,
          warning_count: projectionMetadata.status_summary.warning_count,
        },
        deterministic_identity_summary: {
          normalized_plan_hash:
            projectionMetadata.deterministic_identity_summary.normalized_plan_hash,
          replay_identity:
            projectionMetadata.deterministic_identity_summary.replay_identity,
          validator_run_identity:
            projectionMetadata.deterministic_identity_summary.validator_run_identity,
        },
        semantic_interpretation_summary: {
          scope_resolution:
            projectionMetadata.semantic_interpretation_summary.scope_resolution,
          effective_scope:
            projectionMetadata.semantic_interpretation_summary.effective_scope,
          evidence_source:
            projectionMetadata.semantic_interpretation_summary.evidence_source,
        },
        issues_summary: {
          errors: deepCloneArray(projectionMetadata.issues_summary.errors),
          warnings: deepCloneArray(projectionMetadata.issues_summary.warnings),
        },
      },
      semantic_view: {
        scope_resolution:
          projectionMetadata.semantic_interpretation_summary.scope_resolution,
        effective_scope:
          projectionMetadata.semantic_interpretation_summary.effective_scope,
        evidence_source:
          projectionMetadata.semantic_interpretation_summary.evidence_source,
      },
      issues_view: {
        status: projectionMetadata.status_summary.status,
        error_count: projectionMetadata.status_summary.error_count,
        warning_count: projectionMetadata.status_summary.warning_count,
        errors: deepCloneArray(projectionMetadata.issues_summary.errors),
        warnings: deepCloneArray(projectionMetadata.issues_summary.warnings),
      },
      resolution_view: {
        status: resolutionPreview.status,
        scope_resolution: resolutionPreview.scope_resolution,
        effective_scope: resolutionPreview.effective_scope,
        evidence_source: resolutionPreview.evidence_source,
      },
      rule_evaluation_summary_view: {
        phase_order: [...ruleEvaluationSummaryPreview.phase_order],
        ordered_rules: ruleEvaluationSummaryPreview.ordered_rules.map((rule) => ({
          category: rule.category,
          rule_id: rule.rule_id,
          outcome: rule.outcome,
        })),
      },
      rule_evaluation_trace_view: {
        projection_contract: ruleEvaluationTracePreview.projection_contract,
        projection_kind: ruleEvaluationTracePreview.projection_kind,
        input_artifact: ruleEvaluationTracePreview.input_artifact,
        input_identity: {
          artifact_path: ruleEvaluationTracePreview.input_identity.artifact_path,
          input_fingerprint_sha256:
            ruleEvaluationTracePreview.input_identity.input_fingerprint_sha256,
        },
        deterministic_identity_summary: {
          normalized_plan_hash:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .normalized_plan_hash,
          replay_identity:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .replay_identity,
          validator_run_identity:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .validator_run_identity,
        },
      },
      deterministic_identity_view: {
        projection_metadata: {
          normalized_plan_hash:
            projectionMetadata.deterministic_identity_summary.normalized_plan_hash,
          replay_identity:
            projectionMetadata.deterministic_identity_summary.replay_identity,
          validator_run_identity:
            projectionMetadata.deterministic_identity_summary.validator_run_identity,
        },
        rule_evaluation_trace: {
          normalized_plan_hash:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .normalized_plan_hash,
          replay_identity:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .replay_identity,
          validator_run_identity:
            ruleEvaluationTracePreview.deterministic_identity_summary
              .validator_run_identity,
        },
      },
      raw_payload_debug_view: {
        payload_kind: payloadKind,
        tabfield_count: tabfieldCount,
        tabfield_order: [...tabfieldOrder],
        field_preview: normalizedFieldPreview.map((entry) => ({
          name: entry.name,
          page_property: entry.page_property,
          custom_property: entry.custom_property,
        })),
        projection_metadata: {
          projection_contract: projectionMetadata.projection_contract,
          projection_kind: projectionMetadata.projection_kind,
          input_artifact: projectionMetadata.input_artifact,
          input_identity: {
            artifact_path: projectionMetadata.input_identity.artifact_path,
            input_fingerprint_sha256:
              projectionMetadata.input_identity.input_fingerprint_sha256,
          },
          status_summary: {
            status: projectionMetadata.status_summary.status,
            error_count: projectionMetadata.status_summary.error_count,
            warning_count: projectionMetadata.status_summary.warning_count,
          },
          deterministic_identity_summary: {
            normalized_plan_hash:
              projectionMetadata.deterministic_identity_summary.normalized_plan_hash,
            replay_identity:
              projectionMetadata.deterministic_identity_summary.replay_identity,
            validator_run_identity:
              projectionMetadata.deterministic_identity_summary.validator_run_identity,
          },
          semantic_interpretation_summary: {
            scope_resolution:
              projectionMetadata.semantic_interpretation_summary.scope_resolution,
            effective_scope:
              projectionMetadata.semantic_interpretation_summary.effective_scope,
            evidence_source:
              projectionMetadata.semantic_interpretation_summary.evidence_source,
          },
          issues_summary: {
            errors: deepCloneArray(projectionMetadata.issues_summary.errors),
            warnings: deepCloneArray(projectionMetadata.issues_summary.warnings),
          },
        },
        resolution_preview: {
          status: resolutionPreview.status,
          scope_resolution: resolutionPreview.scope_resolution,
          effective_scope: resolutionPreview.effective_scope,
          evidence_source: resolutionPreview.evidence_source,
        },
        rule_evaluation_summary_preview: {
          phase_order: [...ruleEvaluationSummaryPreview.phase_order],
          ordered_rules: ruleEvaluationSummaryPreview.ordered_rules.map((rule) => ({
            category: rule.category,
            rule_id: rule.rule_id,
            outcome: rule.outcome,
          })),
        },
        rule_evaluation_trace_preview: {
          projection_contract: ruleEvaluationTracePreview.projection_contract,
          projection_kind: ruleEvaluationTracePreview.projection_kind,
          input_artifact: ruleEvaluationTracePreview.input_artifact,
          input_identity: {
            artifact_path: ruleEvaluationTracePreview.input_identity.artifact_path,
            input_fingerprint_sha256:
              ruleEvaluationTracePreview.input_identity.input_fingerprint_sha256,
          },
          deterministic_identity_summary: {
            normalized_plan_hash:
              ruleEvaluationTracePreview.deterministic_identity_summary
                .normalized_plan_hash,
            replay_identity:
              ruleEvaluationTracePreview.deterministic_identity_summary
                .replay_identity,
            validator_run_identity:
              ruleEvaluationTracePreview.deterministic_identity_summary
                .validator_run_identity,
          },
        },
      },
    },
  };
}
