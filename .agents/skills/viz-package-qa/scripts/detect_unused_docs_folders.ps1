Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path "docs")) {
  Write-Host "No docs/ folder present."
  exit 0
}

# Heuristic: folders containing only README.md or .keep are "possibly unused"
$dirs = Get-ChildItem -Directory -Recurse -Force -Path "docs"
$flagged = @()

foreach ($d in $dirs) {
  $items = Get-ChildItem -LiteralPath $d.FullName -Force | Where-Object { -not $_.PSIsContainer }
  if ($items.Count -eq 0) { continue }

  $names = $items.Name
  $nonTrivial = $names | Where-Object { $_ -notin @("README.md", ".keep") }
  if ($nonTrivial.Count -eq 0) {
    $flagged += $d.FullName
  }
}

if ($flagged.Count -gt 0) {
  Write-Host "WARN: docs/ subfolders with only README/.keep (possibly unused):"
  $flagged | ForEach-Object { Write-Host " - $_" }
} else {
  Write-Host "PASS: No obviously empty docs/ subfolders detected by heuristic."
}

exit 0
