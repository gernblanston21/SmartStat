param(
  [Parameter(Mandatory=$false)][string]$RefA = "HEAD~1",
  [Parameter(Mandatory=$false)][string]$RefB = "HEAD"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$git = Get-Command git -ErrorAction SilentlyContinue
if (-not $git) { throw "git not found on PATH" }

git diff --name-only $RefA $RefB
