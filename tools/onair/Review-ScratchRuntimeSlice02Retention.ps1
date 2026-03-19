param(
  [string]$ScratchRoot = "tests/_scratch/runtime-slice-02-readonly-plan-bridge",
  [int]$RetainRecentPassCount = 2,
  [int]$RetainDays = 30,
  [switch]$Delete
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ScratchRoot)) {
  Write-Host "Scratch root not found: $ScratchRoot"
  exit 0
}

$files = Get-ChildItem -LiteralPath $ScratchRoot -Recurse -File
$passTagged = @()
$untagged = @()

foreach ($file in $files) {
  $m = [regex]::Match($file.Name, 'pass(?<num>\d+)_')
  if ($m.Success) {
    $passTagged += [pscustomobject]@{
      Path = $file.FullName
      Name = $file.Name
      PassNumber = [int]$m.Groups['num'].Value
      LastWriteTime = $file.LastWriteTime
    }
  } else {
    $untagged += $file.FullName
  }
}

$maxPass = if ($passTagged.Count -gt 0) {
  ($passTagged | Measure-Object -Property PassNumber -Maximum).Maximum
} else {
  $null
}

$minPassToKeep = if ($null -ne $maxPass) {
  [Math]::Max(0, $maxPass - $RetainRecentPassCount + 1)
} else {
  0
}

$cutoff = (Get-Date).AddDays(-1 * [Math]::Abs($RetainDays))

$candidates = @($passTagged | Where-Object {
  $_.PassNumber -lt $minPassToKeep -and $_.LastWriteTime -lt $cutoff
} | Sort-Object PassNumber, LastWriteTime, Name)

Write-Host "Scratch root: $ScratchRoot"
Write-Host "Total files: $($files.Count)"
Write-Host "Pass-tagged files: $($passTagged.Count)"
Write-Host "Untagged files: $($untagged.Count)"
if ($null -ne $maxPass) {
  Write-Host "Highest pass seen: pass$maxPass"
}
Write-Host "Retain recent pass count: $RetainRecentPassCount"
Write-Host "Retain days cutoff: $RetainDays (older than $($cutoff.ToString('yyyy-MM-dd')) eligible)"
Write-Host "Prune candidates (dry-run): $($candidates.Count)"

if ($untagged.Count -gt 0) {
  Write-Host ""
  Write-Host "Untagged files (manual review; never auto-pruned by this script):"
  $untagged | ForEach-Object { Write-Host "  $_" }
}

if ($candidates.Count -gt 0) {
  Write-Host ""
  Write-Host "Candidate files:"
  $candidates | ForEach-Object {
    Write-Host "  [pass$($_.PassNumber)] $($_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')) $($_.Path)"
  }
}

if ($Delete) {
  Write-Host ""
  Write-Host "Delete switch provided: removing candidate files only."
  foreach ($candidate in $candidates) {
    Remove-Item -LiteralPath $candidate.Path -Force
  }
  Write-Host "Deleted files: $($candidates.Count)"
} else {
  Write-Host ""
  Write-Host "Dry-run only. Use -Delete to remove listed candidate files."
}

