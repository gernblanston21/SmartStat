import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import App, { PANEL_RENDER_ORDER, resolveEvidenceTargets } from "../App";
import { adaptPreviewPayloadToViewModel } from "../adapters/previewPayloadToViewModel";
import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import previewFixture from "../../../../tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json";

function cloneViewModel(viewModel: RuntimePreviewViewModel): RuntimePreviewViewModel {
  return JSON.parse(JSON.stringify(viewModel)) as RuntimePreviewViewModel;
}

function countMatches(haystack: string, pattern: RegExp): number {
  return (haystack.match(pattern) ?? []).length;
}

describe("runtime preview inspector UI scaffold", () => {
  it("renders from frozen fixture via adapter", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    expect(html).toContain("SmartStat Runtime Preview Inspector");
    expect(html).toContain("Read-Only");
    expect(html).toContain("Contract View");
    expect(html).toContain("Semantic vs Resolution");
    expect(html).toContain("Deterministic Identity");
    expect(html).toContain("What this screen is");
    expect(html).toContain("What read-only means");
    expect(html).toContain("What to check next");
    expect(html).toContain("Where Detailed Evidence Lives");
    expect(html).toContain("Consistent Preview");
    expect(html).toContain("Discrepancy Summary");
    expect(html).toContain("Rule Count Snapshot");
    expect(html).toContain("High-priority review items");
    expect(html).toContain("Secondary checks and supporting context");
    expect(html).toContain("All Core Checks PASS");
    expect(html).toContain("Comparison coverage summary");
    expect(html).toContain("Follow links in the last column");
    expect(html).toContain("Panel Order Reference");
    expect(html).toContain("Status key: READY");
    expect(html).toContain("Evidence targets (2)");
    expect(html).toContain('data-evidence-panel-key="semantic_view"');
    expect(html).toContain('data-evidence-panel-key="resolution_view"');
    expect(html).toContain('data-evidence-panel-key="rule_evaluation_summary_view"');
    expect(html).toMatch(/data-evidence-panel-key="semantic_view"[\s\S]*Semantic Interpretation/);
    expect(html).toMatch(
      /data-evidence-panel-key="rule_evaluation_summary_view"[\s\S]*Rule Evaluation Summary/
    );
    expect(html).not.toContain("Mismatch drill-down");
    expect(html).not.toContain('class="secondary-checks" open=""');
    expect(html).toContain('class="comparison-coverage is-pass"');
    expect(html).toContain('data-coverage-count="pass">4</strong>');
    expect(html).toContain('data-coverage-count="mismatch">0</strong>');
    expect(html).toContain('data-coverage-count="unavailable">0</strong>');
    expect(countMatches(html, /data-evidence-panel-key="/g)).toBe(6);
    expect(html).not.toContain("comparison-target-chip is-unavailable");
    expect(html).not.toContain("comparison-evidence-unavailable");
    expect(countMatches(html, /data-panel-status="ready"/g)).toBe(PANEL_RENDER_ORDER.length);
    expect(countMatches(html, /data-panel-jump-status="ready"/g)).toBe(PANEL_RENDER_ORDER.length);
    expect(html).not.toContain('data-panel-status="review"');
    expect(html).not.toContain('data-panel-status="unavailable"');
  });

  it("renders panels in deterministic approved order", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    let previousIndex = -1;
    PANEL_RENDER_ORDER.forEach((panel) => {
      const marker = `data-panel-key="${panel.key}"`;
      const currentIndex = html.indexOf(marker);
      expect(currentIndex).toBeGreaterThan(previousIndex);
      previousIndex = currentIndex;
    });
  });

  it("keeps raw payload in a compact summary with optional full JSON details", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    expect(html).toContain("Field Preview (Contract-Ordered)");
    expect(html).toContain("View full raw payload JSON");
    expect(html).toContain("<details>");
  });

  it("renders side-by-side comparison sections for parity-heavy panels", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    expect(html).toContain("Phase Order Comparison");
    expect(html).toContain("Semantic");
    expect(html).toContain("Resolution");
    expect(html).toContain("Projection Metadata");
    expect(html).toContain("Rule Trace");
    expect(html).toContain("Parity");
  });

  it("renders in-page evidence links that point to detailed panel anchors", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    expect(html).toContain('href="#panel-semantic_view"');
    expect(html).toContain('href="#panel-resolution_view"');
    expect(html).toContain('href="#panel-rule_evaluation_summary_view"');
    expect(html).toContain('id="panel-semantic_view"');
    expect(html).toContain('id="panel-rule_evaluation_summary_view"');
  });

  it("fails closed for invalid evidence target references", () => {
    const resolved = resolveEvidenceTargets([
      "semantic_view",
      "invalid_panel_key",
    ]);

    expect(resolved).toHaveLength(2);
    expect(resolved[0]).toEqual({
      requestedKey: "semantic_view",
      panelKey: "semantic_view",
      title: "Semantic Interpretation",
      href: "#panel-semantic_view",
      isUnavailable: false,
    });
    expect(resolved[1]).toEqual({
      requestedKey: "invalid_panel_key",
      panelKey: null,
      title: "UNAVAILABLE",
      href: null,
      isUnavailable: true,
    });
  });

  it("fails closed for empty evidence target lists", () => {
    const resolved = resolveEvidenceTargets([]);

    expect(resolved).toEqual([
      {
        requestedKey: "__none__",
        panelKey: null,
        title: "No evidence provided in current payload",
        href: null,
        isUnavailable: true,
      },
    ]);
  });

  it("surfaces mismatch state with stronger review emphasis", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const mismatched = cloneViewModel(viewModel);
    mismatched.view_model.semantic_view.scope_resolution = "forced_mismatch";

    const html = renderToStaticMarkup(<App viewModel={mismatched} />);

    expect(html).toContain("Inconsistencies Found");
    expect(html).toContain("Review Needed");
    expect(html).toContain("MISMATCH");
    expect(html).toContain("Mismatch drill-down");
    expect(html).toContain("Semantic vs Resolution - Why this failed");
    expect(html).toContain("SEMANTIC:");
    expect(html).toContain("RESOLUTION:");
    expect(html).toContain("CHK-SEMANTIC-RESOLUTION");
    expect(html).toContain("semantic.scope_resolution = forced_mismatch");
    expect(html).toContain('href="#panel-semantic_view"');
    expect(html).toContain('href="#panel-resolution_view"');
    expect(html).toContain('class="comparison-coverage is-mismatch"');
    expect(html).toContain('data-coverage-count="pass">3</strong>');
    expect(html).toContain('data-coverage-count="mismatch">1</strong>');
    expect(html).toContain('data-coverage-count="unavailable">0</strong>');
    expect(countMatches(html, /data-evidence-panel-key="/g)).toBe(6);
    expect(html).not.toContain("comparison-target-chip is-unavailable");
    expect(html).not.toContain("comparison-evidence-unavailable");
    expect(html).toContain('data-panel-key="semantic_view" data-panel-status="review"');
    expect(html).toContain('data-panel-key="resolution_view" data-panel-status="review"');
    expect(html).toContain('data-panel-jump-key="semantic_view" data-panel-jump-status="review"');
    expect(html).toContain('data-panel-jump-key="resolution_view" data-panel-jump-status="review"');
    expect(html).toContain("Review evidence targeted by CHK-SEMANTIC-RESOLUTION.");
    expect(html).not.toContain("BEFORE:");
    expect(html).not.toContain("AFTER:");
    expect(html).toContain('class="secondary-checks" open=""');
  });

  it("uses expected/actual labels for rule phase order mismatches", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const mismatched = cloneViewModel(viewModel);
    mismatched.view_model.rule_evaluation_summary_view.phase_order = [
      "STRUCTURAL",
      "DETERMINISM",
      "SEMANTIC",
      "BOUNDARY",
    ];

    const html = renderToStaticMarkup(<App viewModel={mismatched} />);

    expect(html).toContain("Rule Phase Order - Why this failed");
    expect(html).toContain("EXPECTED:");
    expect(html).toContain("ACTUAL:");
    expect(html).toContain("CHK-RULE-PHASE-ORDER");
    expect(html).toContain('href="#panel-rule_evaluation_summary_view"');
  });

  it("renders unavailable panel state when a required section is missing", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const missingProjection = cloneViewModel(viewModel);
    (
      missingProjection.view_model as unknown as Record<string, unknown>
    ).projection_metadata_view = undefined;

    const html = renderToStaticMarkup(<App viewModel={missingProjection} />);

    expect(html).toContain("Projection Metadata");
    expect(html).toContain("Unavailable in this preview");
    expect(html).toContain("No evidence provided in current payload");
    expect(html).toContain("UNAVAILABLE");
    expect(html).toContain('class="comparison-coverage is-unavailable"');
    expect(html).toContain('data-coverage-count="pass">3</strong>');
    expect(html).toContain('data-coverage-count="mismatch">0</strong>');
    expect(html).toContain('data-coverage-count="unavailable">1</strong>');
    expect(html).not.toContain("comparison-target-chip is-unavailable");
    expect(html).not.toContain("comparison-evidence-unavailable");
    expect(html).toContain(
      'data-panel-key="projection_metadata_view" data-panel-status="unavailable"'
    );
    expect(html).toContain(
      'data-panel-jump-key="projection_metadata_view" data-panel-jump-status="unavailable"'
    );
    expect(html).not.toContain(
      'data-panel-key="projection_metadata_view" data-panel-status="ready"'
    );
  });

  it("keeps drill-down grounded when some related evidence is unavailable", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const partial = cloneViewModel(viewModel);
    partial.view_model.semantic_view.scope_resolution = "forced_mismatch";
    partial.view_model.resolution_view.evidence_source = "";

    const html = renderToStaticMarkup(<App viewModel={partial} />);

    expect(html).toContain("Semantic vs Resolution - Why this failed");
    expect(html).toContain("Differing grounded fields");
    expect(html).toContain("semantic.scope_resolution = forced_mismatch");
    expect(html).toContain("Some related fields are unavailable in this preview.");
  });

  it("renders calmly when all top comparisons are unavailable", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const unavailable = cloneViewModel(viewModel);

    unavailable.view_model.semantic_view.scope_resolution = "";
    unavailable.view_model.semantic_view.effective_scope = "";
    unavailable.view_model.semantic_view.evidence_source = "";
    unavailable.view_model.resolution_view.scope_resolution = "";
    unavailable.view_model.resolution_view.effective_scope = "";
    unavailable.view_model.resolution_view.evidence_source = "";

    unavailable.view_model.deterministic_identity_view.projection_metadata.normalized_plan_hash = "";
    unavailable.view_model.deterministic_identity_view.projection_metadata.replay_identity = "";
    unavailable.view_model.deterministic_identity_view.projection_metadata.validator_run_identity = "";
    unavailable.view_model.deterministic_identity_view.rule_evaluation_trace.normalized_plan_hash = "";
    unavailable.view_model.deterministic_identity_view.rule_evaluation_trace.replay_identity = "";
    unavailable.view_model.deterministic_identity_view.rule_evaluation_trace.validator_run_identity = "";

    unavailable.view_model.projection_metadata_view.projection_contract = "";
    unavailable.view_model.projection_metadata_view.projection_kind = "";
    unavailable.view_model.projection_metadata_view.input_artifact = "";
    unavailable.view_model.projection_metadata_view.input_identity.artifact_path = "";
    unavailable.view_model.projection_metadata_view.input_identity.input_fingerprint_sha256 = "";

    unavailable.view_model.rule_evaluation_trace_view.projection_contract = "";
    unavailable.view_model.rule_evaluation_trace_view.projection_kind = "";
    unavailable.view_model.rule_evaluation_trace_view.input_artifact = "";
    unavailable.view_model.rule_evaluation_trace_view.input_identity.artifact_path = "";
    unavailable.view_model.rule_evaluation_trace_view.input_identity.input_fingerprint_sha256 = "";

    unavailable.view_model.rule_evaluation_summary_view.phase_order = [];
    unavailable.view_model.rule_evaluation_summary_view.ordered_rules = [];

    const html = renderToStaticMarkup(<App viewModel={unavailable} />);

    expect(html).toContain('class="comparison-coverage is-unavailable"');
    expect(html).toContain('data-coverage-count="pass">0</strong>');
    expect(html).toContain('data-coverage-count="mismatch">0</strong>');
    expect(html).toContain('data-coverage-count="unavailable">4</strong>');
    expect(html).toContain("Mismatch drill-down");
    expect(html).toContain("No evidence provided in current payload");
    expect(html).not.toContain('href="#panel-__none__"');
  });
});
