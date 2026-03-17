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
    expect(html).toContain("High-priority mismatches");
    expect(html).toContain("Secondary checks and supporting context");
    expect(html).toContain("All Core Checks PASS");
    expect(html).toContain("Follow links in the last column");
    expect(html).toContain("Panel Order Reference");
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
    expect(html).toContain('class="secondary-checks" open=""');
  });
});
