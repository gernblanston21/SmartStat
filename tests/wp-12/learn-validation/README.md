# WP-12 Learn Validation Framework

This folder contains WP-12 validation tooling for learn INI safety checks.

## Goal

Verify learn INI files are structurally valid and reject malformed/partial content using tests-only tooling.

## Validator Checks

`validate_learn_ini.ps1` validates:
- parseability (section + key/value line parsing)
- no duplicate keys within a section
- no empty section names or empty keys
- basic key/value format strictness (keys must use `key=value` format)
- partial-write indicators (for example malformed section header, missing `=`, conflict markers)

## RC Safety

In RC mode this validator is read-only and must NOT modify any repo learn files.
It only reads target files and writes artifacts under this folder.

## How To Run

Validate one file:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-12/learn-validation/validate_learn_ini.ps1 -IniPath tests/wp-12/learn-validation/fixtures/good/good_basic.learn.ini
```

Run full WP-12 harness:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-12/learn-validation/run_wp12.ps1 -RunLabel wp12_runA
```

## Artifact Locations

`run_wp12.ps1` writes outputs to:

- `tests/wp-12/learn-validation/artifacts/<runLabel>/cases/*.txt`
- `tests/wp-12/learn-validation/artifacts/<runLabel>/run_summary.txt`
- `tests/wp-12/learn-validation/artifacts/<runLabel>/run_summary.json`

Pass/fail:
- `run_wp12.ps1` exits `0` when all expected outcomes match
- `run_wp12.ps1` exits `2` when any expectation fails
