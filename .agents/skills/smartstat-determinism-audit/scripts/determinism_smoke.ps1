param(
  [Parameter(Mandatory=$false)][string]$ScriptPath = "SmartStat_v4.0.0_beta.vbs",
  [Parameter(Mandatory=$false)][string]$OutDir = "DiagLogs/Harness",
  [Parameter(Mandatory=$false)][string]$Tag = "det_smoke"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $ScriptPath)) { throw "File not found: $ScriptPath" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

function RunOnce([string]$suffix) {
  $out = Join-Path $OutDir ("$Tag" + "_" + $suffix + ".txt")
  # Run via cscript to avoid wscript UI. Capture stdout+stderr.
  $cmd = "cscript.exe //nologo `"$ScriptPath`""
  $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd > `"$out`" 2>&1" -Wait -PassThru
  if ($p.ExitCode -ne 0) {
    Write-Host "WARN: cscript exit code: $($p.ExitCode). Output: $out"
  }
  return $out
}

$a = RunOnce "A"
Start-Sleep -Milliseconds 150
$b = RunOnce "B"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$hashScript = Join-Path $here "stable_hash_diag.ps1"

$ha = & $hashScript -Path $a
$hb = & $hashScript -Path $b

Write-Host "A: $($ha.Hash)  $a"
Write-Host "B: $($hb.Hash)  $b"

if ($ha.Hash -ne $hb.Hash) {
  Write-Host "FAIL: Determinism smoke mismatch. Diffing raw outputs:"
  $diffScript = Join-Path $here "diff_text.ps1"
  & $diffScript -A $a -B $b
  exit 2
}

Write-Host "PASS: Determinism smoke hashes match."
exit 0
