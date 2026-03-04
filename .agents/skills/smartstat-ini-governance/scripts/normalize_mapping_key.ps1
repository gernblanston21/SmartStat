param(
  [Parameter(Mandatory=$true)][string]$Value
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# SmartStat normalization rule: replace spaces with underscores, collapse multiple underscores.
$out = $Value.Trim()
$out = $out -replace '\s+', '_'
$out = $out -replace '_{2,}', '_'
Write-Output $out
