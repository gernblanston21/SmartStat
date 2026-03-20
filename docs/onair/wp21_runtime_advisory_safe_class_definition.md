# WP21 Runtime Advisory Safe Class Definition

Status date: `2026-03-20`  
Scope: `Docs-only governance class definition`  
Class name: `RUNTIME_ADVISORY_SAFE_CLASS`

## Purpose

Define a narrow governance class for runtime-adjacent operator-facing advisory
features that are safe to develop without granting runtime execution authority.

This class exists to separate:

- runtime context awareness
from
- execution authority

## Core Properties

`RUNTIME_ADVISORY_SAFE_CLASS` must always be:

1. read-only
2. advisory-only
3. non-authoritative
4. non-mutating
5. deterministic
6. fail-closed
7. reversible

## Allowed Uses

Allowed uses are limited to bounded advisory awareness work such as:

1. operator-facing advisory checkpoint support
2. advisory message/summary surfaces with zero action authority
3. non-runtime tooling that reads existing approved artifacts and emits
   advisory-only output
4. workflow/process integration artifacts that improve operator decision clarity
5. bounded `.vbs` read-only advisory logic only when all class constraints are
   satisfied

## Bounded `.vbs` Read-Only Advisory Scope

Allowed `.vbs` advisory logic in this class must remain:

1. read-only and advisory-only
2. deterministic and fail-closed
3. non-authoritative (never execution permission)
4. isolated from apply/take/cue paths
5. isolated from tabfield write paths
6. isolated from socket execution paths
7. non-mutating to runtime payloads and output semantics

## Forbidden Uses

The following are always forbidden inside this class:

1. runtime/apply/take/cue execution behavior
2. mutation of runtime state or runtime payloads
3. tabfield writes
4. socket execution behavior
5. any implicit or explicit execution permission
6. viewer truth-surface expansion
7. semantic reinterpretation of runtime truth surfaces
8. any `.vbs` advisory logic that changes SmartStat output semantics
9. any `.vbs` advisory logic that can trigger execution behavior indirectly

## Safety Rule: Never Execution Permission

Advisory status or advisory visibility must never be interpreted as permission
to execute operator actions. Operator judgment remains the final authority.

## Relationship to Existing Runtime Gates

This class does not:

1. authorize broad WP-20 runtime re-entry
2. authorize broad `.vbs` runtime implementation by itself
3. authorize mutation/apply execution paths

WP-20 freeze/gate protections remain fully in force.

## Implementation Policy Under This Class

Codex may proceed with bounded implementation work without a new runtime
re-entry envelope each time only when all class boundaries are explicitly
satisfied.

If any boundary is ambiguous, or any forbidden surface may be touched, work must
halt and move to a separate explicit bounded authorization envelope.

For `.vbs` scope specifically:

- only bounded read-only advisory logic is in class scope
- any uncertainty about execution coupling requires fail-closed halt and
  envelope escalation

## References

- `AGENTS.md`
- `docs/onair/wp20_runtime_reentry_authorization_envelope_01.md`
- `docs/onair/wp21_narrow_runtime_advisory_reentry_envelope_16.md`
- `docs/onair/viewer-readonly-boundary.md`
- `docs/onair/wp21_decision_layer_definition.md`
- `docs/viz-trio/environment_constraints.md`
