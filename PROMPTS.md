# PROMPTS.md — SmartStat / Codex Prompt Library

This file is a copy/paste library for consistent Codex runs.
It complements (does not replace) AGENTS.md governance.

---

## P0 — Universal Codex Kickoff (Viz Trio doc-gated)

You MUST treat /docs/viz-trio as the source of truth for Viz Trio behavior.

Before proposing ANY SmartStat changes:
1) Read docs/viz-trio/environment_constraints.md (live-safe rules)
2) Read docs/viz-trio/command_reference.md and docs/viz-trio/commands_full_index.md (command assumptions)
3) Read docs/viz-trio/tabfields.md (tabfield prefix heuristics)
4) Use docs/viz-trio/page_list.md, page_editor.md, show_control.md when reasoning about operator workflow

If any proposed change conflicts with those docs, STOP and propose a safer alternative.
When unsure, QUOTE the relevant doc section and justify the decision.
Fail closed.

Now proceed with the task below.

---

## P1 — Resume Roadmap on v4_Dev (post-RC baseline)

Resume Post-RC Roadmap on v4_Dev.

BASELINE:
- v4_Dev contains RC1 freeze merged back in.
- RC1 behavior is the baseline; do not re-review RC1 unless explicitly requested.

RULES:
- No naming convention changes.
- No INI key reordering.
- No placeholders.
- Provide unified diffs (minimum -U5).
- Flag SmartStatTrayApp compatibility risks.
- Preserve fail-closed gating semantics.

TASK:
[PASTE WP TASK HERE]

---

## P2 — WP Kickoff Template (fill-in)

Start Work Package: [WP-##] — [TITLE]

SCOPE:
- [What this WP changes]
- [What it explicitly does NOT change]

DEFINITION OF DONE:
- [Bullet list]

DELIVERABLES:
- [Files touched]
- [Validation artifacts]
- [Changelog entry]

CONSTRAINTS:
- Preserve determinism unless WP explicitly changes it
- Fail closed on ambiguity / validation gates
- No refactors outside scope

REQUEST:
1) Identify the exact code locations to change (file + function/sub name).
2) Propose the smallest safe implementation.
3) Provide unified diffs (-U5) with context.
4) Provide regression impact + validation steps.

---

## P3 — Read-Only Audit (no file changes)

READ-ONLY AUDIT MODE

RULES:
- DO NOT MODIFY ANY FILES.
- No formatting, no whitespace changes, no refactors.
- Only inspect and report.

OUTPUT:
- Markdown report with file + line anchors.
- Code excerpts <= 10 lines per excerpt.

AUDIT TARGET:
[DESCRIBE TARGET]

AUDIT QUESTIONS:
- [Question 1]
- [Question 2]
- [Question 3]

---

## P4 — Diff Review Request (for patches Codex proposes)

Review the proposed diff for:
- Determinism drift
- Ambiguity leakage (fail-open paths)
- INI key order/format violations
- External tool compatibility risks (TrayApp)
- Hidden refactors outside scope

Return:
- Approve / reject
- Specific required changes (by hunk)
- Any tests/validation needed

Paste diff below:
[DIFF]
