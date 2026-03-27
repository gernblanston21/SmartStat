param(
  [Parameter(Mandatory=$true)][string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) { throw "File not found: $Path" }

$lines = Get-Content -LiteralPath $Path
$section = "<ROOT>"
$keys = @{}
$fail = $false

function Flush([string]$sec, [hashtable]$kmap) {
  foreach ($k in $kmap.Keys) {
    $v = $kmap[$k]
    if ($v.Count -gt 1) {
      Write-Host "FAIL: Duplicate key '$k' in section $sec at lines: $($v -join ', ')"
      $script:fail = $true
    }
  }
}

for ($i=0; $i -lt $lines.Count; $i++) {
  $raw = $lines[$i]
  $lineNo = $i + 1
  $line = $raw.Trim()

  if ($line -match '^\s*\[.*\]\s*$') {
    Flush $section $keys
    $section = $line
    $keys = @{}
    continue
  }

  if ($line -eq "" -or $line.StartsWith(";") -or $line.StartsWith("'") -or $line.StartsWith("#")) {
    continue
  }

  if ($line -match '^\s*([^=]+?)\s*=\s*(.*)\s*$') {
    $key = $Matches[1].Trim()
    if (-not $keys.ContainsKey($key)) { $keys[$key] = New-Object System.Collections.Generic.List[int] }
    $null = $keys[$key].Add($lineNo)
  }
}

Flush $section $keys

if ($fail) { exit 2 }
Write-Host "PASS: No duplicate keys found: $Path"
exit 0
