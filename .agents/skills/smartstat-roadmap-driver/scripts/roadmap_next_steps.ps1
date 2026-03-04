param(
  [Parameter(Mandatory=$false)][string]$RoadmapPath = "ROADMAP.md",
  [Parameter(Mandatory=$false)][string]$SessionPath = "SESSION.md"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ReadOrNull([string]$p) {
  if (Test-Path $p) { return Get-Content -LiteralPath $p -Raw }
  return $null
}

function InferMode([string]$roadmap, [string]$session) {
  if ($session) {
    $sessionHasActiveRc = (
      $session -match '(?im)RC1 Stabilization Discipline \(Active\)' -or
      $session -match '(?im)\bv4\.1\.0_RC1\b' -or
      $session -match '(?im)Stabilization Discipline \(Active\)'
    )
    if ($sessionHasActiveRc) { return "RC" }

    $sessionHasExplicitRcInactiveOrComplete = (
      $session -match '(?im)\bRC\s*discipline\s+is\s+inactive\b' -or
      $session -match '(?im)\bRC\s*discipline\s+is\s+complete\b' -or
      $session -match '(?im)\bStabilization Discipline\s*\((Inactive|Complete)\)\b'
    )
    if ($sessionHasExplicitRcInactiveOrComplete -and -not $sessionHasActiveRc) { return "DEV" }
  }
  if ($roadmap -and $roadmap -match '(?ims)##\s*Current State\s*(.+?)(\r?\n\r?\n|$)') {
    $blk = $Matches[1]
    if ($blk -match '(?im)\bRC1\b' -or $blk -match '(?im)\b_RC\b' -or $blk -match '(?im)\bstabilization\b') { return "RC" }
  }
  return "DEV"
}

$roadmap = ReadOrNull $RoadmapPath
$session = ReadOrNull $SessionPath
if (-not $roadmap) { throw "Missing required file: $RoadmapPath" }

$mode = InferMode $roadmap $session
Write-Host "MODE=$mode"
Write-Host ""

if ($mode -eq "RC") {
  Write-Host "RC_ALLOWED_NEXT_STEPS:"
  Write-Host " - Run STRICT harness repeat-run verification and archive evidence under tests/."
  Write-Host " - Improve log clarity (no logic changes)."
  Write-Host " - Add/standardize regression packs (WP-11) under tests/."
  Write-Host " - Update docs/viz-trio if grounding gaps are found."
  Write-Host ""
  Write-Host "NOT_ALLOWED (RC): behavior/logic changes unless explicitly approved as a roadmap item with evidence."
  exit 0
}

Write-Host "DEV_NEXT_STEPS:"
Write-Host " - WP-11: Harness regression pack framework (define pack + standard compare rules)."
Write-Host " - WP-12: Enhanced learn system validation (validate writes; quarantine bad writes)."
Write-Host " - WP-13: Performance optimization with no behavioral diffs."
Write-Host " - WP-14: TrayApp alignment preparation (contract + compatibility checks)."
Write-Host ""
Write-Host "TIP: Use wp_prompt.ps1 to generate a Codex prompt for a chosen WP/target."
exit 0
