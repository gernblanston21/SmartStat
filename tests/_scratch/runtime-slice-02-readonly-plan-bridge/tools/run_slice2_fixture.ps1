param(
  [Parameter(Mandatory = $true)]
  [string]$Fixture,
  [string]$RunLabel = "slice2_run"
)

$fixturePath = (Resolve-Path $Fixture).Path
$outDir = "tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs"
$evidenceOut = Join-Path $outDir ($RunLabel + ".json")
$diagOut = Join-Path $outDir ($RunLabel + ".log")

& cscript //nologo SmartStat_v4.1.0.vbs "/slice_gate:WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE" "/runtime_fixture:$fixturePath" "/evidence_out:$evidenceOut" "/diag_out:$diagOut"
if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

$result = Get-Content -Raw $evidenceOut | ConvertFrom-Json
Write-Output ("status=" + $result.status)
Write-Output ("error_code=" + $result.error_code)
Write-Output ("evidence_out=" + $evidenceOut)
Write-Output ("diag_out=" + $diagOut)
