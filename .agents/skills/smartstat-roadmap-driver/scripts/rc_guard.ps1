param(
  [Parameter(Mandatory=$true)][string]$Mode,
  [Parameter(Mandatory=$true)][string]$ChangeDescription
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$modeNorm = $Mode.Trim().ToUpperInvariant()
$desc = $ChangeDescription.Trim()
$lower = $desc.ToLowerInvariant()

if ($modeNorm -ne "RC") {
  Write-Host "PASS: Mode is not RC."
  exit 0
}

# Fail-closed heuristic: if it sounds like behavior/logic changes and NOT clearly tests/docs/log-only.
$behaviorSignals = @(
  "logic","behavior","resolver","resolution","precedence","mapping","output","syntax",
  "qualifier","category","ambiguity","gate","transaction","commit","parser","refactor",
  "implement","support","feature","optimize","performance"
)
$allowedSignals = @("docs","documentation","comment","log clarity","logging","tests","test","harness","regression","evidence")

$looksBehavior = $false
foreach ($w in $behaviorSignals) { if ($lower -like "*$w*") { $looksBehavior = $true } }

$looksAllowed = $false
foreach ($w in $allowedSignals) { if ($lower -like "*$w*") { $looksAllowed = $true } }

if ($looksBehavior -and -not $looksAllowed) {
  Write-Host "FAIL: RC mode request appears behavior-affecting and not clearly docs/tests/log-only."
  Write-Host "DESC=$desc"
  exit 2
}

Write-Host "PASS: RC guard did not detect a behavior-only change (or it is framed as tests/docs/log work)."
exit 0
