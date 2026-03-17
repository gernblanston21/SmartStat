import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import App, { PANEL_RENDER_ORDER } from "../App";
import { adaptPreviewPayloadToViewModel } from "../adapters/previewPayloadToViewModel";
import { RuntimePreviewViewModel } from "../contracts/runtimePreviewIntake";
import previewFixture from "../../../../tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json";

function cloneViewModel(viewModel: RuntimePreviewViewModel): RuntimePreviewViewModel {
  return JSON.parse(JSON.stringify(viewModel)) as RuntimePreviewViewModel;
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
    expect(html).toContain("Follow links in the last column");
    expect(html).toContain("Panel Order Reference");
    expect(html).not.toContain("Mismatch drill-down");
    expect(html).not.toContain('class="secondary-checks" open=""');
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
});
