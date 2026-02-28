# SmartStat v4.0.0_RC1 --- Release Candidate

Release Type: Stability / Production Readiness Base Version: v4.0.0_beta
Branch: v4_RC Date: (fill at tagging)

------------------------------------------------------------------------

# Overview

v4.0.0_RC1 is a production readiness candidate built on the frozen
`v4.0.0_beta` baseline.

No runtime semantics were modified.

This release focuses on logging clarity, defensive hardening, harness
verification, and documentation refinement.

------------------------------------------------------------------------

# Included Changes

## LOG

## (List logging-only changes here)

-   
-   

## GUARD

## (List defensive changes here)

-   
-   

## HARNESS

## (List harness improvements here)

-   
-   

## DOC

## (List documentation updates here)

-   
-   

## HYGIENE

## (List repo-only cleanup here)

-   
-   

------------------------------------------------------------------------

# Behavioral Verification

The following guarantees from v4.0.0_beta remain intact:

-   Fail-closed ambiguity enforcement.
-   FIRST_SEEN deterministic resolver tie rule.
-   Strict transaction validation.
-   STRICT harness commit gating.
-   Hard failure for OUTMAP.EMPTY.
-   No silent early exits.
-   No unbounded ambiguity dumps.

No scoring math, thresholds, selection ordering, or schema behavior was
altered.

------------------------------------------------------------------------

# Validation Summary

-   STRICT harness executed across representative templates.
-   Zero-diff runs allowed commit.
-   Non-zero-diff runs blocked commit.
-   No regression in ambiguity gating.
-   No unexpected runtime errors.

------------------------------------------------------------------------

# Known Limitations

-   Resolver candidate enumeration remains dependent on dictionary key
    order (UNSORTED_ENUM).
-   Tie rule remains FIRST_SEEN by design in v4.0.x line.
-   Deterministic enumeration refactor deferred to v4.1.0.

------------------------------------------------------------------------

# Upgrade Notes

No configuration schema changes. No INI changes required. No
SmartStatTrayApp contract changes.

This release is backward-compatible with v4.0.0_beta configuration.

------------------------------------------------------------------------

# Tagging

git tag v4.0.0_RC1 git push origin v4.0.0_RC1

------------------------------------------------------------------------

End of changelog.
