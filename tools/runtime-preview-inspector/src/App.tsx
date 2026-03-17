import {
  EXPECTED_RULE_PHASE_ORDER,
  RulePhase,
  RuntimePreviewViewModel,
} from "./contracts/runtimePreviewIntake";
import { DeterministicIdentityPanel } from "./panels/DeterministicIdentityPanel";
import { IntakeHeaderPanel } from "./panels/IntakeHeaderPanel";
import { IssuesPanel } from "./panels/IssuesPanel";
import { ProjectionMetadataPanel } from "./panels/ProjectionMetadataPanel";
import { RawPayloadPanel } from "./panels/RawPayloadPanel";
import { ResolutionPanel } from "./panels/ResolutionPanel";
import { RuleSummaryPanel } from "./panels/RuleSummaryPanel";
import { RuleTracePanel } from "./panels/RuleTracePanel";
import { SemanticPanel } from "./panels/SemanticPanel";

interface AppProps {
  viewModel: RuntimePreviewViewModel;
}

export const PANEL_RENDER_ORDER = [
  { key: "intake_header", title: "Intake Header" },
  { key: "projection_metadata_view", title: "Projection Metadata" },
  { key: "semantic_view", title: "Semantic Interpretation" },
  { key: "issues_view", title: "Issues Summary" },
  { key: "resolution_view", title: "Resolution Preview" },
  { key: "rule_evaluation_summary_view", title: "Rule Evaluation Summary" },
  { key: "rule_evaluation_trace_view", title: "Rule Evaluation Trace" },
  { key: "deterministic_identity_view", title: "Deterministic Identity" },
  { key: "raw_payload_debug_view", title: "Raw Payload Debug" },
] as const;

type PanelKey = (typeof PANEL_RENDER_ORDER)[number]["key"];

const PANEL_TITLE_BY_KEY = Object.fromEntries(
  PANEL_RENDER_ORDER.map((panel) => [panel.key, panel.title])
) as Record<PanelKey, string>;

function parityLabel(value: boolean): "PASS" | "MISMATCH" {
  return value ? "PASS" : "MISMATCH";
}

function comparisonPairingKey(key: ComparisonCheckKey): string {
  return `CHK-${key.replace(/_/g, "-").toUpperCase()}`;
}

interface InspectionSignal {
  key: string;
  label: string;
  pass: boolean;
  detail: string;
  priority: "High" | "Medium";
}

type ComparisonCheckKey =
  | "semantic_resolution"
  | "deterministic_identity"
  | "traceability_overlap"
  | "rule_phase_order";

interface ComparisonCheck {
  key: ComparisonCheckKey;
  label: string;
  pass: boolean;
  priority: "High" | "Medium";
  pairingKey: string;
  leftTag: string;
  rightTag: string;
  evidenceTargets: PanelKey[];
}

interface DrillDownDiffLine {
  key: string;
  expectedLabel: string;
  expectedValue: string;
  actualLabel: string;
  actualValue: string;
  mismatch: boolean;
}

interface DrillDownCard {
  key: ComparisonCheckKey;
  label: string;
  priority: "High" | "Medium";
  pairingKey: string;
  leftTag: string;
  rightTag: string;
  evidenceTargets: PanelKey[];
  contributingFields: string[];
  diffLines: DrillDownDiffLine[];
}

function formatDiffValue(value: string | readonly string[]): string {
  if (typeof value === "string") {
    return value;
  }
  return `[${value.join(", ")}]`;
}

function createDiffLine(
  key: string,
  expectedLabel: string,
  expectedValue: string | readonly string[],
  actualLabel: string,
  actualValue: string | readonly string[]
): DrillDownDiffLine {
  const expected = formatDiffValue(expectedValue);
  const actual = formatDiffValue(actualValue);
  return {
    key,
    expectedLabel,
    expectedValue: expected,
    actualLabel,
    actualValue: actual,
    mismatch: expected !== actual,
  };
}

export default function App({ viewModel }: AppProps): JSX.Element {
  const sections = viewModel.view_model;
  const phaseOrderParity =
    JSON.stringify(sections.rule_evaluation_summary_view.phase_order) ===
    JSON.stringify(EXPECTED_RULE_PHASE_ORDER);
  const semanticResolutionParity =
    sections.semantic_view.scope_resolution === sections.resolution_view.scope_resolution &&
    sections.semantic_view.effective_scope === sections.resolution_view.effective_scope &&
    sections.semantic_view.evidence_source === sections.resolution_view.evidence_source;
  const deterministicIdentityParity =
    sections.deterministic_identity_view.projection_metadata.normalized_plan_hash ===
      sections.deterministic_identity_view.rule_evaluation_trace.normalized_plan_hash &&
    sections.deterministic_identity_view.projection_metadata.replay_identity ===
      sections.deterministic_identity_view.rule_evaluation_trace.replay_identity &&
      sections.deterministic_identity_view.projection_metadata.validator_run_identity ===
        sections.deterministic_identity_view.rule_evaluation_trace.validator_run_identity;
  const traceabilityOverlapParity =
    sections.projection_metadata_view.projection_contract ===
      sections.rule_evaluation_trace_view.projection_contract &&
    sections.projection_metadata_view.projection_kind ===
      sections.rule_evaluation_trace_view.projection_kind &&
    sections.projection_metadata_view.input_artifact ===
      sections.rule_evaluation_trace_view.input_artifact &&
    sections.projection_metadata_view.input_identity.artifact_path ===
      sections.rule_evaluation_trace_view.input_identity.artifact_path &&
      sections.projection_metadata_view.input_identity.input_fingerprint_sha256 ===
      sections.rule_evaluation_trace_view.input_identity.input_fingerprint_sha256;
  const topStatusPass = sections.issues_view.status === "PASS";
  const hasNoErrors = sections.issues_view.error_count === 0;
  const hasNoWarnings = sections.issues_view.warning_count === 0;

  const ruleCountsByPhase = sections.rule_evaluation_summary_view.ordered_rules.reduce<
    Record<RulePhase, number>
  >(
    (acc, rule) => {
      acc[rule.category] += 1;
      return acc;
    },
    {
      STRUCTURAL: 0,
      SEMANTIC: 0,
      DETERMINISM: 0,
      BOUNDARY: 0,
    }
  );

  const inspectionSignals: InspectionSignal[] = [
    {
      key: "preview_status",
      label: "Preview status",
      pass: topStatusPass,
      detail: sections.issues_view.status,
      priority: "High",
    },
    {
      key: "error_count",
      label: "Error count",
      pass: hasNoErrors,
      detail: String(sections.issues_view.error_count),
      priority: "High",
    },
    {
      key: "semantic_resolution",
      label: "Semantic vs resolution",
      pass: semanticResolutionParity,
      detail: semanticResolutionParity ? "All key fields match" : "Scope/evidence mismatch",
      priority: "High",
    },
    {
      key: "deterministic_identity",
      label: "Deterministic identity parity",
      pass: deterministicIdentityParity,
      detail: deterministicIdentityParity ? "All identity fields match" : "Identity mismatch",
      priority: "High",
    },
    {
      key: "traceability_overlap",
      label: "Projection vs trace overlap",
      pass: traceabilityOverlapParity,
      detail: traceabilityOverlapParity ? "Overlap fields aligned" : "Overlap mismatch",
      priority: "Medium",
    },
    {
      key: "rule_phase_order",
      label: "Rule phase order",
      pass: phaseOrderParity,
      detail: phaseOrderParity ? "Expected deterministic order" : "Order differs from contract",
      priority: "Medium",
    },
    {
      key: "warning_count",
      label: "Warning count",
      pass: hasNoWarnings,
      detail: String(sections.issues_view.warning_count),
      priority: "Medium",
    },
  ];
  const mismatchSignals = inspectionSignals.filter((signal) => !signal.pass);
  const passSignals = inspectionSignals.length - mismatchSignals.length;
  const highPrioritySignals = inspectionSignals.filter((signal) => signal.priority === "High");
  const highPriorityMismatches = highPrioritySignals.filter((signal) => !signal.pass);
  const mismatchCount = mismatchSignals.length;
  const highPriorityMismatchCount = highPriorityMismatches.length;
  const overviewHealthy = mismatchCount === 0;
  const reviewTone = overviewHealthy ? "Consistent Preview" : "Inconsistencies Found";
  const reviewToneClass = overviewHealthy ? "is-pass" : "is-mismatch";
  const topStatusToneLabel = overviewHealthy ? "All Core Checks PASS" : "Review Needed";
  const comparisonChecks: ComparisonCheck[] = [
    {
      key: "semantic_resolution",
      label: "Semantic vs Resolution",
      pass: semanticResolutionParity,
      priority: "High",
      pairingKey: comparisonPairingKey("semantic_resolution"),
      leftTag: "SEMANTIC",
      rightTag: "RESOLUTION",
      evidenceTargets: ["semantic_view", "resolution_view"],
    },
    {
      key: "deterministic_identity",
      label: "Deterministic Identity",
      pass: deterministicIdentityParity,
      priority: "High",
      pairingKey: comparisonPairingKey("deterministic_identity"),
      leftTag: "PROJECTION",
      rightTag: "RULE TRACE",
      evidenceTargets: ["deterministic_identity_view"],
    },
    {
      key: "traceability_overlap",
      label: "Traceability Overlap",
      pass: traceabilityOverlapParity,
      priority: "Medium",
      pairingKey: comparisonPairingKey("traceability_overlap"),
      leftTag: "PROJECTION",
      rightTag: "RULE TRACE",
      evidenceTargets: ["projection_metadata_view", "rule_evaluation_trace_view"],
    },
    {
      key: "rule_phase_order",
      label: "Rule Phase Order",
      pass: phaseOrderParity,
      priority: "Medium",
      pairingKey: comparisonPairingKey("rule_phase_order"),
      leftTag: "EXPECTED",
      rightTag: "ACTUAL",
      evidenceTargets: ["rule_evaluation_summary_view"],
    },
  ];
  const ruleCategorySequence = Array.from(
    new Set(sections.rule_evaluation_summary_view.ordered_rules.map((rule) => rule.category))
  );
  const drillDownDetailsByKey: Record<
    ComparisonCheckKey,
    Pick<DrillDownCard, "contributingFields" | "diffLines">
  > = {
    semantic_resolution: {
      contributingFields: [
        "semantic.scope_resolution",
        "semantic.effective_scope",
        "semantic.evidence_source",
        "resolution.scope_resolution",
        "resolution.effective_scope",
        "resolution.evidence_source",
      ],
      diffLines: [
        createDiffLine(
          "scope_resolution",
          "semantic.scope_resolution",
          sections.semantic_view.scope_resolution,
          "resolution.scope_resolution",
          sections.resolution_view.scope_resolution
        ),
        createDiffLine(
          "effective_scope",
          "semantic.effective_scope",
          sections.semantic_view.effective_scope,
          "resolution.effective_scope",
          sections.resolution_view.effective_scope
        ),
        createDiffLine(
          "evidence_source",
          "semantic.evidence_source",
          sections.semantic_view.evidence_source,
          "resolution.evidence_source",
          sections.resolution_view.evidence_source
        ),
      ],
    },
    deterministic_identity: {
      contributingFields: [
        "projection_metadata.normalized_plan_hash",
        "projection_metadata.replay_identity",
        "projection_metadata.validator_run_identity",
        "rule_trace.normalized_plan_hash",
        "rule_trace.replay_identity",
        "rule_trace.validator_run_identity",
      ],
      diffLines: [
        createDiffLine(
          "normalized_plan_hash",
          "projection_metadata.normalized_plan_hash",
          sections.deterministic_identity_view.projection_metadata.normalized_plan_hash,
          "rule_trace.normalized_plan_hash",
          sections.deterministic_identity_view.rule_evaluation_trace.normalized_plan_hash
        ),
        createDiffLine(
          "replay_identity",
          "projection_metadata.replay_identity",
          sections.deterministic_identity_view.projection_metadata.replay_identity,
          "rule_trace.replay_identity",
          sections.deterministic_identity_view.rule_evaluation_trace.replay_identity
        ),
        createDiffLine(
          "validator_run_identity",
          "projection_metadata.validator_run_identity",
          sections.deterministic_identity_view.projection_metadata.validator_run_identity,
          "rule_trace.validator_run_identity",
          sections.deterministic_identity_view.rule_evaluation_trace.validator_run_identity
        ),
      ],
    },
    traceability_overlap: {
      contributingFields: [
        "projection_contract",
        "projection_kind",
        "input_artifact",
        "input_identity.artifact_path",
        "input_identity.input_fingerprint_sha256",
      ],
      diffLines: [
        createDiffLine(
          "projection_contract",
          "projection_metadata.projection_contract",
          sections.projection_metadata_view.projection_contract,
          "rule_trace.projection_contract",
          sections.rule_evaluation_trace_view.projection_contract
        ),
        createDiffLine(
          "projection_kind",
          "projection_metadata.projection_kind",
          sections.projection_metadata_view.projection_kind,
          "rule_trace.projection_kind",
          sections.rule_evaluation_trace_view.projection_kind
        ),
        createDiffLine(
          "input_artifact",
          "projection_metadata.input_artifact",
          sections.projection_metadata_view.input_artifact,
          "rule_trace.input_artifact",
          sections.rule_evaluation_trace_view.input_artifact
        ),
        createDiffLine(
          "artifact_path",
          "projection_metadata.input_identity.artifact_path",
          sections.projection_metadata_view.input_identity.artifact_path,
          "rule_trace.input_identity.artifact_path",
          sections.rule_evaluation_trace_view.input_identity.artifact_path
        ),
        createDiffLine(
          "input_fingerprint_sha256",
          "projection_metadata.input_identity.input_fingerprint_sha256",
          sections.projection_metadata_view.input_identity.input_fingerprint_sha256,
          "rule_trace.input_identity.input_fingerprint_sha256",
          sections.rule_evaluation_trace_view.input_identity.input_fingerprint_sha256
        ),
      ],
    },
    rule_phase_order: {
      contributingFields: [
        "rule_evaluation_summary.phase_order",
        "ordered_rules[*].category",
      ],
      diffLines: [
        createDiffLine(
          "phase_order",
          "expected.phase_order",
          EXPECTED_RULE_PHASE_ORDER,
          "actual.phase_order",
          sections.rule_evaluation_summary_view.phase_order
        ),
        createDiffLine(
          "rule_category_sequence",
          "expected.category_sequence",
          EXPECTED_RULE_PHASE_ORDER,
          "actual.category_sequence",
          ruleCategorySequence
        ),
      ],
    },
  };
  const mismatchDrillDownCards: DrillDownCard[] = comparisonChecks
    .filter((check) => !check.pass)
    .map((check) => {
      const details = drillDownDetailsByKey[check.key];
      return {
        ...check,
        contributingFields: details.contributingFields,
        diffLines: details.diffLines.filter((diffLine) => diffLine.mismatch),
      };
    });

  const orderedPanels: Array<{
    key: PanelKey;
    className: string;
    node: JSX.Element;
  }> =
    [
      {
        key: "intake_header",
        className: "panel-slot panel-slot--half",
        node: <IntakeHeaderPanel view={sections.intake_header} />,
      },
      {
        key: "projection_metadata_view",
        className: "panel-slot panel-slot--half",
        node: (
          <ProjectionMetadataPanel
            view={sections.projection_metadata_view}
            traceView={sections.rule_evaluation_trace_view}
          />
        ),
      },
      {
        key: "semantic_view",
        className: "panel-slot panel-slot--half",
        node: (
          <SemanticPanel
            view={sections.semantic_view}
            resolutionView={sections.resolution_view}
          />
        ),
      },
      {
        key: "issues_view",
        className: "panel-slot panel-slot--half",
        node: <IssuesPanel view={sections.issues_view} />,
      },
      {
        key: "resolution_view",
        className: "panel-slot panel-slot--half",
        node: <ResolutionPanel view={sections.resolution_view} />,
      },
      {
        key: "rule_evaluation_summary_view",
        className: "panel-slot panel-slot--full",
        node: <RuleSummaryPanel view={sections.rule_evaluation_summary_view} />,
      },
      {
        key: "rule_evaluation_trace_view",
        className: "panel-slot panel-slot--half",
        node: <RuleTracePanel view={sections.rule_evaluation_trace_view} />,
      },
      {
        key: "deterministic_identity_view",
        className: "panel-slot panel-slot--half",
        node: <DeterministicIdentityPanel view={sections.deterministic_identity_view} />,
      },
      {
        key: "raw_payload_debug_view",
        className: "panel-slot panel-slot--full",
        node: <RawPayloadPanel view={sections.raw_payload_debug_view} />,
      },
    ];

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>SmartStat Runtime Preview Inspector</h1>
        <p className="app-summary">
          This screen helps you review one saved preview file. It is read-only:
          nothing here writes to graphics systems, runs runtime actions, or changes
          SmartStat state.
        </p>
        <ul className="status-tags">
          <li>Read-Only</li>
          <li>Preview</li>
          <li>Contract View</li>
          <li>Traceability</li>
          <li className={`status-tag--tone ${reviewToneClass}`}>{topStatusToneLabel}</li>
        </ul>
        <section className="overview-strip" aria-label="Screen overview">
          <article className="overview-card">
            <h3>What this screen is</h3>
            <p>
              A viewer for one frozen preview payload that shows whether important
              sections agree with each other.
            </p>
          </article>
          <article className="overview-card">
            <h3>What read-only means</h3>
            <p>
              You can inspect values, but you cannot run, apply, refresh, or mutate
              anything from this page.
            </p>
          </article>
          <article className="overview-card">
            <h3>How to review quickly</h3>
            <p>
              Start with high-priority checks, then use secondary context and panel
              detail only when something looks off.
            </p>
          </article>
        </section>
        <section className={`review-hero ${reviewToneClass}`} aria-label="Review outcome summary">
          <div className="review-hero-main">
            <h2>{reviewTone}</h2>
            <p>
              PASS checks: <strong>{passSignals}</strong> / {inspectionSignals.length}. MISMATCH
              checks: <strong>{mismatchCount}</strong>.
            </p>
            <h3>Discrepancy Summary</h3>
            {mismatchCount ? (
              <ul className="compact-list">
                {mismatchSignals.map((signal) => (
                  <li key={`mismatch-${signal.key}`}>
                    <strong>{signal.label}</strong>: {signal.detail}
                  </li>
                ))}
              </ul>
            ) : (
              <p className="panel-note">
                No discrepancy signals detected in this preview.
              </p>
            )}
          </div>
          <div className="review-hero-metrics" aria-label="Review metrics">
            <article className={`metric-chip ${highPriorityMismatchCount > 0 ? "is-mismatch" : "is-muted"}`}>
              <span>High-priority mismatches</span>
              <strong>{highPriorityMismatchCount}</strong>
            </article>
            <article className={`metric-chip ${mismatchCount > 0 ? "is-mismatch" : "is-muted"}`}>
              <span>Total mismatches</span>
              <strong>{mismatchCount}</strong>
            </article>
            <article className={`metric-chip ${hasNoErrors ? "is-muted" : "is-mismatch"}`}>
              <span>Error count</span>
              <strong>{sections.issues_view.error_count}</strong>
            </article>
            <article className={`metric-chip ${hasNoWarnings ? "is-muted" : "is-mismatch"}`}>
              <span>Warning count</span>
              <strong>{sections.issues_view.warning_count}</strong>
            </article>
          </div>
        </section>
        <section className="check-first-strip" aria-label="What to check next">
          <h2>What to check next</h2>
          <p className={`check-first-result ${reviewToneClass}`}>{reviewTone}</p>
          <p className="check-first-note">
            Follow links in the last column to jump directly to detailed evidence panels.
          </p>
          <table className="summary-table summary-table--priority">
            <thead>
              <tr>
                <th>Comparison</th>
                <th>Result</th>
                <th>Where Detailed Evidence Lives</th>
              </tr>
            </thead>
            <tbody>
              {comparisonChecks.map((check) => (
                <tr
                  key={check.key}
                  className={`priority-row priority-${check.priority.toLowerCase()} ${
                    check.pass ? "is-pass" : "is-mismatch"
                  }`}
                >
                  <td>
                    <div className="comparison-label-cell">
                      <strong>{check.label}</strong>
                      <div className="comparison-cell-meta">
                        <span
                          className={`comparison-key-chip priority-${check.priority.toLowerCase()}`}
                        >
                          {check.pairingKey}
                        </span>
                      </div>
                    </div>
                  </td>
                  <td>
                    <span className={`parity-pill ${check.pass ? "is-pass" : "is-mismatch"}`}>
                      {parityLabel(check.pass)}
                    </span>
                  </td>
                  <td>
                    <ul className="comparison-evidence-list">
                      {check.evidenceTargets.map((panelKey) => (
                        <li key={`${check.key}:${panelKey}`}>
                          <a href={`#panel-${panelKey}`}>{PANEL_TITLE_BY_KEY[panelKey]}</a>
                        </li>
                      ))}
                    </ul>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="check-first-hint">
            High-priority mismatches: <strong>{highPriorityMismatchCount}</strong>. Total
            mismatches: <strong>{mismatchCount}</strong>.
          </p>
          {mismatchDrillDownCards.length > 0 ? (
            <section className="drilldown-strip" aria-label="Mismatch drill-down">
              <h3>Mismatch drill-down</h3>
              <p className="panel-note">
                Each card summarizes what failed, how values differ, and where detailed
                evidence lives.
              </p>
              <div className="drilldown-list">
                {mismatchDrillDownCards.map((card) => (
                  <article
                    key={`drilldown-${card.key}`}
                    className={`drilldown-card ${
                      card.priority === "High" ? "is-high" : "is-medium"
                    }`}
                  >
                    <div className="drilldown-card-header">
                      <h4>{card.label} - Why this failed</h4>
                      <div className="drilldown-card-meta">
                        <span
                          className={`comparison-key-chip priority-${card.priority.toLowerCase()}`}
                        >
                          {card.pairingKey}
                        </span>
                        <span className="drilldown-priority">{card.priority} Priority</span>
                      </div>
                    </div>
                    <p className="drilldown-diff-count">
                      Differing fields: <strong>{card.diffLines.length}</strong>
                    </p>
                    <ul className="drilldown-diff-list">
                      {card.diffLines.map((diffLine) => (
                        <li key={`diff-${card.key}-${diffLine.key}`} className="drilldown-diff-row">
                          <p className="drilldown-field-name">{diffLine.key}</p>
                          <div>
                            <span className="diff-label diff-label-before">{card.leftTag}:</span>{" "}
                            <code>{`${diffLine.expectedLabel} = ${diffLine.expectedValue}`}</code>
                          </div>
                          <div>
                            <span className="diff-label diff-label-after">{card.rightTag}:</span>{" "}
                            <code>{`${diffLine.actualLabel} = ${diffLine.actualValue}`}</code>
                          </div>
                        </li>
                      ))}
                    </ul>
                    <p className="drilldown-fields-title">Key contributing fields</p>
                    <ul className="drilldown-fields">
                      {card.contributingFields.map((fieldName) => (
                        <li key={`field-${card.key}-${fieldName}`}>
                          <code>{fieldName}</code>
                        </li>
                      ))}
                    </ul>
                    <p className="drilldown-links-title">Related evidence panels:</p>
                    <ul className="comparison-evidence-list">
                      {card.evidenceTargets.map((panelKey) => (
                        <li key={`drilldown-link-${card.key}-${panelKey}`}>
                          <a href={`#panel-${panelKey}`}>{PANEL_TITLE_BY_KEY[panelKey]}</a>
                        </li>
                      ))}
                    </ul>
                  </article>
                ))}
              </div>
            </section>
          ) : null}
        </section>
        <details className="secondary-checks" open={!overviewHealthy}>
          <summary>Secondary checks and supporting context</summary>
          <p className="panel-note">
            Use this section after the high-priority checks. It provides additional
            parity and count context without crowding the first scan.
          </p>
          <p className="panel-note">
            If a top comparison fails, inspect related panels below in deterministic order.
            This context helps explain why a mismatch may exist.
          </p>
          <nav className="panel-jump-nav" aria-label="Panel order reference">
            <h3>Panel Order Reference</h3>
            <ol className="panel-jump-list">
              {PANEL_RENDER_ORDER.map((panel) => (
                <li key={`jump-${panel.key}`}>
                  <a href={`#panel-${panel.key}`}>{panel.title}</a>
                </li>
              ))}
            </ol>
          </nav>
          <section className="secondary-grid">
            <article className="overview-card">
              <h3>Rule Count Snapshot</h3>
              <table className="summary-table">
                <thead>
                  <tr>
                    <th>Phase</th>
                    <th>Rules</th>
                  </tr>
                </thead>
                <tbody>
                  {EXPECTED_RULE_PHASE_ORDER.map((phase) => (
                    <tr key={`phase-count-${phase}`}>
                      <td>{phase}</td>
                      <td>{ruleCountsByPhase[phase]}</td>
                    </tr>
                  ))}
                  <tr>
                    <td>Total Ordered Rules</td>
                    <td>{sections.rule_evaluation_summary_view.ordered_rules.length}</td>
                  </tr>
                </tbody>
              </table>
            </article>
            <article className="overview-card">
              <h3>Traceability Snapshot</h3>
              <dl className="kv-grid compact-kv-grid">
                <dt>Projection Contract</dt>
                <dd>{sections.projection_metadata_view.projection_contract}</dd>
                <dt>Projection Kind</dt>
                <dd>{sections.projection_metadata_view.projection_kind}</dd>
                <dt>Input Artifact</dt>
                <dd>{sections.projection_metadata_view.input_artifact}</dd>
              </dl>
            </article>
          </section>
        </details>
      </header>

      <main className="panel-stack" aria-label="Runtime Preview Panels">
        {orderedPanels.map((panel) => (
          <section
            key={panel.key}
            id={`panel-${panel.key}`}
            data-panel-key={panel.key}
            className={panel.className}
          >
            {panel.node}
          </section>
        ))}
      </main>
    </div>
  );
}
