param(
  [Parameter(Mandatory=$true)][string]$WP,
  [Parameter(Mandatory=$false)][string]$Target = "",
  [Parameter(Mandatory=$false)][string]$Mode = "DEV"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$wpNorm = $WP.Trim()
$targetNorm = $Target.Trim()
$modeNorm = $Mode.Trim().ToUpperInvariant()

$task = "Continue SmartStat roadmap work for $wpNorm"
if ($targetNorm) { $task += " — $targetNorm" }

Write-Output "Tell Codex This (copy/paste as-is):"
Write-Output ""
Write-Output "You are working in the SmartStat repo."
Write-Output "Obey AGENTS.md governance, Viz Trio grounding (docs/viz-trio/), and determinism/RC discipline."
Write-Output ""
Write-Output "MODE: $modeNorm"
Write-Output "TASK: $task"
Write-Output ""
Write-Output "Deliverables:"
Write-Output "- Unified diff (-U5)"
Write-Output "- Regression impact summary"
Write-Output "- Validation steps"
Write-Output "- Determinism evidence (hashes + diff if applicable)"
Write-Output ""
Write-Output "If RC mode and the task is behavior-affecting, fail closed and propose an RC-allowed alternative."
Write-Output ""
