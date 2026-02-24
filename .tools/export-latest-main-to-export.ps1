param(
  [string]$WorkspaceFolder
)

$srcDir = Join-Path $WorkspaceFolder "Main_TrioScript"
$outDir = Join-Path $WorkspaceFolder "_EXPORT"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$f = Get-ChildItem $srcDir -Filter '*.vbs' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $f) {
  Write-Host "No .vbs found in Main_TrioScript"
  exit 1
}

$out = Join-Path $outDir "SmartStat_TRIO_PASTE.vbs"
Get-Content $f.FullName -Raw | Set-Content -Path $out -Encoding Default

Write-Host ("Exported: " + $out)
Write-Host ("Source:   " + $f.Name)
