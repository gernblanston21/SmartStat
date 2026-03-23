# MLB Learn-Layer Controlled Proposal Summary (PASS_06)

## Scope

- Pass: `MLB_LEARN_LAYER_CONTROLLED_PROPOSAL_PASS_06`
- Inputs: PASS_05 staging artifacts from `tests/_scratch/onair-mlb-pass05-learn-layer-staging/`.
- Mode: proposal-only; no SmartStat file mutation.
- No `.vbs` changes, no `.ini` changes, no runtime integration.

## Learn-File Structure Findings

- Reference file: `SmartStat_Mappings.learn.ini`
- Observed pending section order:
  - `PENDING`
  - `PENDING_ALIASES`
  - `PENDING_QUALIFIER_ALIASES`
  - `PENDING_QUALIFIER`
  - `PENDING_ALIASES_PITCHER`
- Formatting conventions followed:
  - section headers as `[SECTION]`
  - entry lines as `KEY=VALUE` (no spaces around `=`)
  - comment lines prefixed with `;`
- Proposal mirrors pending-section structure only (candidate blocks), not a synthetic full-file rewrite.

## Proposal Findings

- Safe-to-stage input candidates: **27**
- Promoted into INI-shaped proposal body: **2**
- Dropped from safe_to_stage due concrete-value/fit constraints: **25**
- Proposal grouping by section:
  - `PENDING`: 0
  - `PENDING_ALIASES`: 2
  - `PENDING_QUALIFIER_ALIASES`: 0
  - `PENDING_QUALIFIER`: 0
  - `PENDING_ALIASES_PITCHER`: 0
- Dropped safe-to-stage rationale: candidates with `TBD_CANONICAL_TARGET` were excluded fail-closed from proposal body.

## Exclusions

- `manual_review_required`: 280
- `composite_defer`: 238
- `rejected_or_not_yet_useful`: 480
- These buckets remained out of proposal insertion by policy (safe_to_stage-only promotion).

## Governance Boundary

- No runtime changes
- No INI changes
- No automatic integration

## Recommended Next Pass

- `MLB_LEARN_LAYER_PATCH_AUTHORIZATION_PASS_07`
- Scope: review and explicitly authorize (or reject) each proposed PASS_06 INI entry before any controlled INI edit pass.
