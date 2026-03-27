param(
  [Parameter(Mandatory=$true)][string]$A,
  [Parameter(Mandatory=$true)][string]$B
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $A)) { throw "File not found: $A" }
if (-not (Test-Path $B)) { throw "File not found: $B" }

# Use git diff if available for nicer output; fallback to Compare-Object
$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
  git diff --no-index -- $A $B
  exit $LASTEXITCODE
}

$la = Get-Content -LiteralPath $A
$lb = Get-Content -LiteralPath $B
Compare-Object -ReferenceObject $la -DifferenceObject $lb -IncludeEqual:$false | Format-Table -AutoSize
exit 0
