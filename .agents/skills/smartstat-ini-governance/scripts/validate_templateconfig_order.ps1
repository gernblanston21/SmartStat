param(
  [Parameter(Mandatory=$false)][string]$Path = "SmartStat_TemplateConfig.ini"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) { throw "File not found: $Path" }

# Required order for TemplateConfig sections (when keys are present)
$order = @(
  "config_id",
  "entity",
  "qualifier",
  "filter_tabfields",
  "category_tabfields",
  "row_limit",
  "output_map"
)

$lines = Get-Content -LiteralPath $Path
$section = ""
$seen = @{}
$fail = $false

function Flush-Section([string]$sec, [hashtable]$keys) {
  if (-not $sec) { return }
  if ($sec -notmatch '^\[TEMPLATE:') { return }

  # Build the encountered sequence (only keys that are in $order)
  $encountered = @()
  foreach ($k in $keys.Keys) {
    if ($order -contains $k) { $encountered += $k }
  }

  # Validate relative ordering: for any pair in $order, indices must be ascending if both present.
  for ($i = 0; $i -lt $order.Count; $i++) {
    for ($j = $i + 1; $j -lt $order.Count; $j++) {
      $a = $order[$i]
      $b = $order[$j]
      if ($keys.ContainsKey($a) -and $keys.ContainsKey($b)) {
        if ($keys[$a].FirstLine -gt $keys[$b].FirstLine) {
          Write-Host "FAIL: $sec ordering violation: '$a' appears after '$b' (line $($keys[$a].FirstLine) > $($keys[$b].FirstLine))"
          $script:fail = $true
        }
      }
    }
  }

  # Qualifier must exist if it's a TEMPLATE section (hard rule for SmartStat generation discipline)
  if (-not $keys.ContainsKey("qualifier")) {
    Write-Host "FAIL: $sec missing required key: qualifier"
    $script:fail = $true
  }
}

for ($i=0; $i -lt $lines.Count; $i++) {
  $raw = $lines[$i]
  $lineNo = $i + 1
  $line = $raw.Trim()

  if ($line -match '^\s*\[.*\]\s*$') {
    Flush-Section $section $seen
    $section = $line
    $seen = @{}
    continue
  }

  if ($line -eq "" -or $line.StartsWith(";") -or $line.StartsWith("'") -or $line.StartsWith("#")) {
    continue
  }

  if ($line -match '^\s*([^=]+?)\s*=\s*(.*)\s*$') {
    $key = $Matches[1].Trim()
    if (-not $seen.ContainsKey($key)) {
      $seen[$key] = [pscustomobject]@{ FirstLine = $lineNo; Raw = $raw }
    }
    continue
  }
}

Flush-Section $section $seen

if ($fail) {
  exit 2
} else {
  Write-Host "PASS: TemplateConfig ordering + qualifier presence validated: $Path"
  exit 0
}
