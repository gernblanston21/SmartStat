param(
  [Parameter(Mandatory=$false)][string]$RoadmapPath = "ROADMAP.md",
  [Parameter(Mandatory=$false)][string]$SessionPath = "SESSION.md",
  [Parameter(Mandatory=$false)][string]$AgentsPath  = "AGENTS.md"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ReadOrNull([string]$p) {
  if (Test-Path $p) { return Get-Content -LiteralPath $p -Raw }
  return $null
}

function InferMode([string]$roadmap, [string]$session) {
  $mode = "DEV"
  if ($session) {
    $sessionHasActiveRc = (
      $session -match '(?im)RC1 Stabilization Discipline \(Active\)' -or
      $session -match '(?im)\bv4\.1\.0_RC1\b' -or
      $session -match '(?im)Stabilization Discipline \(Active\)'
    )
    if ($sessionHasActiveRc) {
      return "RC"
    }

    $sessionHasExplicitRcInactiveOrComplete = (
      $session -match '(?im)\bRC\s*discipline\s+is\s+inactive\b' -or
      $session -match '(?im)\bRC\s*discipline\s+is\s+complete\b' -or
      $session -match '(?im)\bStabilization Discipline\s*\((Inactive|Complete)\)\b'
    )
    if ($sessionHasExplicitRcInactiveOrComplete -and -not $sessionHasActiveRc) {
      return "DEV"
    }
  }
  if ($roadmap) {
    # Look specifically in the Current State block first
    if ($roadmap -match '(?ims)##\s*Current State\s*(.+?)(\r?\n\r?\n|$)') {
      $blk = $Matches[1]
      if ($blk -match '(?im)\bRC1\b' -or $blk -match '(?im)\b_RC\b' -or $blk -match '(?im)\bstabilization\b') {
        return "RC"
      }
    }
    # fallback scan
    if ($roadmap -match '(?im)\bv4\.0\.0_RC\b' -or $roadmap -match '(?im)\bRC1\b') {
      return "RC"
    }
  }
  return $mode
}

$roadmap = ReadOrNull $RoadmapPath
$session = ReadOrNull $SessionPath
$agents  = ReadOrNull $AgentsPath

$missing = @()
if (-not $roadmap) { $missing += $RoadmapPath }
if (-not $agents)  { $missing += $AgentsPath }

if ($missing.Count -gt 0) {
  Write-Host "MISSING_FILES:"
  $missing | ForEach-Object { Write-Host " - $_" }
  exit 2
}

$mode = InferMode $roadmap $session
Write-Host "MODE=$mode"
Write-Host ""

# Extract WP headings: your roadmap uses "## WP-11 (v4.1.0): ..."
$lines = $roadmap -split "`r?`n"
$wpList = New-Object System.Collections.Generic.List[object]

for ($i=0; $i -lt $lines.Count; $i++) {
  $ln = $lines[$i]
  if ($ln -match '^\s*##\s*(WP-\d{1,3})\b(.*)$') {
    $wp = $Matches[1].Trim()
    $titleRest = if (($Matches.Count -ge 3) -and ($null -ne $Matches[2])) { $Matches[2].Trim() } else { "" }
    $title = ($wp + " " + $titleRest).Trim()

    # Section body extends until the next WP heading.
    $j = $i + 1
    while ($j -lt $lines.Count -and ($lines[$j] -notmatch '^\s*##\s*WP-\d{1,3}\b')) {
      $j++
    }

    # Status rule:
    # 1) CLOSED token in heading
    # 2) Else CLOSED evidence line in section body: "^<WP> CLOSED"
    # 3) Else OPEN
    $closedInHeadingRx = '^\s*##\s*' + [regex]::Escape($wp) + '\b.*\bCLOSED\b'
    $closedEvidenceRx = '^\s*' + [regex]::Escape($wp) + '\s+CLOSED\b'
    $hasClosedEvidence = $false
    if ($ln -notmatch $closedInHeadingRx) {
      for ($k = $i + 1; $k -lt $j; $k++) {
        if ($lines[$k] -imatch $closedEvidenceRx) {
          $hasClosedEvidence = $true
          break
        }
      }
    }
    $status = if (($ln -imatch $closedInHeadingRx) -or $hasClosedEvidence) { "CLOSED" } else { "OPEN" }

    $wpList.Add([pscustomobject]@{
      WP = $wp
      Status = $status
      Title = $title
    }) | Out-Null
  }
}

if ($wpList.Count -eq 0) {
  Write-Host "WARN: No WP headings detected with pattern '## WP-##'."
  exit 0
}

Write-Host "WP_COUNT=$($wpList.Count)"
Write-Host ""
$wpList | Format-Table -AutoSize

exit 0
