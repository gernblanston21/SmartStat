import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import App, { PANEL_RENDER_ORDER } from "../App";
import { adaptPreviewPayloadToViewModel } from "../adapters/previewPayloadToViewModel";
import previewFixture from "../../../../tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json";

describe("runtime preview inspector UI scaffold", () => {
  it("renders from frozen fixture via adapter", () => {
    const viewModel = adaptPreviewPayloadToViewModel(previewFixture);
    const html = renderToStaticMarkup(<App viewModel={viewModel} />);

    expect(html).toContain("SmartStat Runtime Preview Inspector");
    expect(html).toContain("Read-Only");
    expect(html).toContain("Contract View");
    expect(html).toContain("Contract parity checks");
    expect(html).toContain("Semantic vs Resolution");
    expect(html).toContain("Deterministic Identity");
    expect(html).toContain("What this screen is");
    expect(html).toContain("What read-only means");
    expect(html).toContain("What to check first");
    expect(html).toContain("Consistent Preview");
    expect(html).toContain("Discrepancy Summary");
    expect(html).toContain("Rule Count Snapshot");
    expect(html).toContain("High-priority mismatches");
    expect(html).toContain("Secondary checks and supporting context");
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
});
