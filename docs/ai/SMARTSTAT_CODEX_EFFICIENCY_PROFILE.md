# 🧠 SmartStat Codex Efficiency Profile

## 🎯 Core Principle
Only pay for context when the task actually requires it.

Start with the cheapest safe mode and escalate only when necessary.

---

## ⚙️ Mode System

### 1. Daily Driver Mode

**Settings**
- Auto-context: Minimized
- File scope: Single explicit file or selected code
- Model: Standard

**Use for**
- VBScript edits
- INI fixes
- Bug fixes
- Small enhancements

**Rules**
- Do not refactor unrelated code
- Preserve structure
- Fail closed
- Output only bounded change

---

### 2. Precision Patch Mode (PRIMARY MODE)

**Settings**
- Auto-context: Minimized
- Context: Function + 20–40 surrounding lines OR tagged file

**Use for**
- Runtime fixes
- Mapping logic
- Output_map fixes
- Core engine updates

**Rules**
- Only modify supplied surface
- No scope expansion without audit

---

### 3. Read-Only Audit Mode

**Settings**
- Multi-file context allowed
- No code changes

**Use for**
- Determinism audits
- Regression analysis
- Scope discovery
- Pre-change planning

**Rules**
- Identify exact change scope
- No implementation

---

### 4. Cross-File Implementation Mode

**Settings**
- Multi-file context allowed (ONLY after audit)
- Explicit file list

**Use for**
- Proven multi-file changes

**Rules**
- Only modify approved files
- No opportunistic refactors

---

### 5. Closeout Mode

**Use for**
- Finalizing approved passes

**Rules**
- Commit only approved files
- Update SESSION.md
- Update ROADMAP.md if needed
- Triage untracked files
- Keep `_scratch` artifacts untracked
- DO NOT start next pass

---

### 6. New Slice Definition Mode

**Use for**
- Runtime lane advancement
- WP decisions

**Rules**
- Repo truth first:
  - AGENTS.md
  - SESSION.md
  - ROADMAP.md
  - WP docs
- Propose ONE slice or NO-GO
- No implementation

---

## 🔁 Escalation Rule

Escalate modes ONLY if:
- cross-file dependency uncertainty
- repo-truth verification required
- architecture boundary unclear
- planning new slice
- validation requires multi-file analysis

---

## 🚫 Hard Workflow Rules

- Never send no-op prompts
- Every prompt must:
  - implement
  - audit
  - validate
  - close out
  - triage

- Do not use broad context unless required

---

## ⚡ Cost Controls

- Prefer tagged files over repo-wide context
- Prefer function blocks over full files
- Audit before cross-file work
- Use high reasoning only when necessary

---

## 🧠 Mental Model

| Mode | Cost | Purpose |
|------|------|--------|
| Daily Driver | Low | Surgical edits |
| Precision Patch | Low-Med | Targeted logic |
| Audit | Med | Planning/analysis |
| Cross-File | Med-High | Coordinated edits |
| Slice Definition | Med-High | Architecture decisions |
| Closeout | Low | Finalization |

---

## 🏁 Bottom Line

You don’t want Codex to see everything.

You want it to see exactly enough.
