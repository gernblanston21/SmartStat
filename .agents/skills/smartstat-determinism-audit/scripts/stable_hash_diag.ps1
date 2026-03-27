param(
  [Parameter(Mandatory=$true)][string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) { throw "File not found: $Path" }

$content = Get-Content -LiteralPath $Path -Raw

# Normalize common volatility:
# - Leading timestamp like [YYYY-MM-DD HH:MM:SS]
# - RUN_ID=...
# - MACHINE=... USER=...
# - Any absolute Windows paths (coarse)
$norm = $content
$norm = $norm -replace '\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]\s*', '[TS] '
$norm = $norm -replace 'RUN_ID=\d{8}_\d{6}', 'RUN_ID=[RUN]'
$norm = $norm -replace 'MACHINE=.*? USER=.*?(\r?\n)', 'MACHINE=[M] USER=[U]$1'
$norm = $norm -replace '[A-Z]:\\[^ \r\n]+', '[PATH]'

# Hash normalized content (SHA256)
$bytes = [System.Text.Encoding]::UTF8.GetBytes($norm)
$sha = [System.Security.Cryptography.SHA256]::Create()
$hashBytes = $sha.ComputeHash($bytes)
$hash = ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ""

[pscustomobject]@{
  Path = $Path
  Hash = $hash
  NormalizedLength = $norm.Length
}
