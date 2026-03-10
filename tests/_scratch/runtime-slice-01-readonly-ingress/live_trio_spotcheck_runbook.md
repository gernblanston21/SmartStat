# Live Viz Trio Spot-Check Runbook (Operator-Controlled)

Scope: `WP20_RUNTIME_SLICE_01_READONLY_INGRESS`  
Runtime line: `SmartStat_v4.1.0.vbs`  
Pass type: validation/evidence only  

## Explicit Non-Claim

This runbook describes **operator-controlled live validation** steps.  
Codex did **not** execute a live Viz Trio proof in this pass.

## Preconditions

1. `SmartStat_v4.0.0_beta.vbs` and `SmartStat_v4.1.0.vbs` are both available in the live environment.
2. Operator has a known, approved trigger to execute each runtime script in Trio.
3. Two equivalent pages are prepared from the same template with equivalent starting values.
4. Evidence directory exists:
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/`

## Invocation Safety Note (Fail-Closed)

Unsupported by docs - failing closed:
- Local `docs/viz-trio/*` grounding provides command surfaces for read operations and script command categories.
- It does **not** define one universal command that invokes your deployed SmartStat runtime binding in every environment.
- Therefore, use your existing operator-approved baseline/successor execution trigger in live Trio.

## Gate-OFF Parity Spot-Check Method

Goal: confirm `v4.1.0` gate-OFF runtime behavior parity with frozen baseline.

1. On `PAGE_A` (baseline path), execute baseline runtime using operator-approved trigger.
2. On `PAGE_B` (successor path), execute successor runtime with gate OFF (default, no slice gate).
3. Capture read-only snapshots from each page using Trio command line:
   - `page:read <PAGE_A_NR>`
   - `page:getpagename`
   - `page:getpagetemplate`
   - `page:get_tabfield_names`
   - For each returned tabfield `<TAB>`:
     - `page:get_property <TAB>`
     - `tabfield:get_custom_property <TAB>`
4. Save PAGE_A captures to:
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateoff_baseline_snapshot.txt`
5. Repeat step 3 for `PAGE_B`, save to:
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateoff_successor_snapshot.txt`
6. Normalize and hash compare snapshots:

```powershell
$a = Get-Content tests\_scratch\runtime-slice-01-readonly-ingress\runs\live_gateoff_baseline_snapshot.txt
$b = Get-Content tests\_scratch\runtime-slice-01-readonly-ingress\runs\live_gateoff_successor_snapshot.txt
$aNorm = ($a | ForEach-Object { $_.Trim().ToUpper() } | Sort-Object) -join "`n"
$bNorm = ($b | ForEach-Object { $_.Trim().ToUpper() } | Sort-Object) -join "`n"
$sha = [System.Security.Cryptography.SHA256]::Create()
$aHash = [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($aNorm))).Replace('-','')
$bHash = [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($bNorm))).Replace('-','')
Write-Output "live_gateoff_baseline_sha256=$aHash"
Write-Output "live_gateoff_successor_sha256=$bHash"
Write-Output "live_gateoff_hash_match=$([bool]($aHash -eq $bHash))"
```

Expected result: `live_gateoff_hash_match=True`.

## Gate-ON Ingress-Only Spot-Check Method

Goal: confirm gate-ON read-only ingress behaves deterministically and fail-closed.

1. On a chosen page, execute successor runtime with slice-1 gate ON using the operator-approved successor trigger path that passes:
   - `slice_gate=WP20_RUNTIME_SLICE_01_READONLY_INGRESS`
2. Capture produced evidence files:
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run1.json`
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run1.log`
3. Repeat on same inputs:
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run2.json`
   - `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run2.log`
4. Hash compare deterministic outputs:

```powershell
$h1 = (Get-FileHash tests\_scratch\runtime-slice-01-readonly-ingress\runs\live_gateon_ingress_run1.json -Algorithm SHA256).Hash
$h2 = (Get-FileHash tests\_scratch\runtime-slice-01-readonly-ingress\runs\live_gateon_ingress_run2.json -Algorithm SHA256).Hash
Write-Output "live_gateon_run1_sha256=$h1"
Write-Output "live_gateon_run2_sha256=$h2"
Write-Output "live_gateon_hash_match=$([bool]($h1 -eq $h2))"
```

Expected result: `live_gateon_hash_match=True`.

## Required Live Evidence Files

1. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateoff_baseline_snapshot.txt`
2. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateoff_successor_snapshot.txt`
3. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run1.json`
4. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run1.log`
5. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run2.json`
6. `tests/_scratch/runtime-slice-01-readonly-ingress/runs/live_gateon_ingress_run2.log`
