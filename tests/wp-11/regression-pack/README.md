# WP-11 Regression Pack Framework

This folder defines a repeatable harness regression pack for SmartStat.

Scope:
- Tests/framework only (RC-safe)
- No SmartStat core code changes
- No INI schema/order changes

## Files

- `pack_manifest.json`: case catalog (runner + inputs + artifact name)
- `run_pack.ps1`: executes all enabled cases and captures artifacts
- `compare_pack.ps1`: stable-hash comparison between two runs (+ diff on mismatch)
- `artifacts/`: run outputs (created by scripts)

## Add A Case

1. Open `pack_manifest.json`.
2. Add a new object under `cases`:
   - `id`: unique stable case id
   - `enabled`: true/false
   - `runner`: VBScript harness runner path
   - `script` (optional): SmartStat script path override
   - `args`: arguments passed after `<scriptPath>`
   - `artifact`: output file name under run `cases/`
3. If you need parity checks within the same run, add an entry to `within_run_equal`.

## Run The Pack

From repo root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/run_pack.ps1 -RunLabel runA
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/run_pack.ps1 -RunLabel runB
```

Artifacts are written to:

- `tests/wp-11/regression-pack/artifacts/<runLabel>/cases/*.txt`
- `tests/wp-11/regression-pack/artifacts/<runLabel>/run_summary.json`

`run_pack.ps1` exits:
- `0` when all case runs return exit code `0`
- `2` when one or more cases fail

## Compare Two Runs

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/compare_pack.ps1 -RunA runA -RunB runB
```

Compare outputs:

- Stable SHA256 hashes of normalized case artifacts
- Mismatch diffs (if any) under:
  - `tests/wp-11/regression-pack/artifacts/compare/<runA>__<runB>/diffs/`

Reports:

- `tests/wp-11/regression-pack/artifacts/compare/<runA>__<runB>/compare_report.txt`
- `tests/wp-11/regression-pack/artifacts/compare/<runA>__<runB>/compare_report.json`

`compare_pack.ps1` exits:
- `0` when all checks pass
- `2` when any check fails

## Evidence

Passing example commands:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/run_pack.ps1 -RunLabel wp11_runA
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/run_pack.ps1 -RunLabel wp11_runB
powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-11/regression-pack/compare_pack.ps1 -RunA wp11_runA -RunB wp11_runB
```

Passing run artifact locations:

- `tests/wp-11/regression-pack/artifacts/wp11_runA/`
- `tests/wp-11/regression-pack/artifacts/wp11_runB/`
- `tests/wp-11/regression-pack/artifacts/compare/wp11_runA__wp11_runB/`
- `tests/wp-11/regression-pack/artifacts/compare/wp11_runA__wp11_runB/compare_report.txt`

Terminal output style (compare report path):

```text
COMPARE_ROOT=E:\EDRIVE\UNIVERSAL\SmartStat\tests\wp-11\regression-pack\artifacts\compare\wp11_runA__wp11_runB
REPORT_TXT=E:\EDRIVE\UNIVERSAL\SmartStat\tests\wp-11\regression-pack\artifacts\compare\wp11_runA__wp11_runB\compare_report.txt
```
