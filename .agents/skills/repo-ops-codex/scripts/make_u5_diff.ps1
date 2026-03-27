param(
  [Parameter(Mandatory=$false)][string]$RefA = "HEAD~1",
  [Parameter(Mandatory=$false)][string]$RefB = "HEAD",
  [Parameter(Mandatory=$false)][string]$Path = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$git = Get-Command git -ErrorAction SilentlyContinue
if (-not $git) { throw "git not found on PATH" }

if ($Path -and (Test-Path $Path)) {
  git diff -U5 $RefA $RefB -- $Path
  exit $LASTEXITCODE
}

git diff -U5 $RefA $RefB
exit $LASTEXITCODE
