param(
  [Parameter(Mandatory=$true)][string]$Path,
  [Parameter(Mandatory=$true)][int]$Start,
  [Parameter(Mandatory=$true)][int]$End
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) { throw "File not found: $Path" }
if ($Start -lt 1 -or $End -lt 1 -or $End -lt $Start) { throw "Invalid range: $Start-$End" }

$lines = Get-Content -LiteralPath $Path
$max = $lines.Count
if ($Start -gt $max) { throw "Start out of bounds. File has $max lines." }
if ($End -gt $max) { $End = $max }

for ($i=$Start; $i -le $End; $i++) {
  $idx = $i - 1
  "{0,6}: {1}" -f $i, $lines[$idx]
}
