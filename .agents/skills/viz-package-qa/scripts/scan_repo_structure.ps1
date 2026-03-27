Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$required = @(
  "docs",
  "governance",
  "tests",
  "SmartStat_v4.0.0_beta.vbs",
  "SmartStat_TemplateConfig.ini",
  "SmartStat_Mappings.ini",
  "SmartStat_StaticOverrides.ini"
)

$missing = @()
foreach ($p in $required) {
  if (-not (Test-Path $p)) { $missing += $p }
}

if ($missing.Count -gt 0) {
  Write-Host "FAIL: Missing required repo paths:"
  $missing | ForEach-Object { Write-Host " - $_" }
  exit 2
}

Write-Host "PASS: Required structure present."

# Identify empty directories (excluding .git and .vscode)
$empties = Get-ChildItem -Directory -Recurse -Force |
  Where-Object { $_.FullName -notmatch '\\\.git(\\|$)' -and $_.FullName -notmatch '\\\.vscode(\\|$)' } |
  Where-Object {
    $items = Get-ChildItem -LiteralPath $_.FullName -Force
    $items.Count -eq 0
  }

if ($empties.Count -gt 0) {
  Write-Host ""
  Write-Host "WARN: Empty directories detected:"
  $empties | ForEach-Object { Write-Host " - $($_.FullName)" }
}

exit 0
