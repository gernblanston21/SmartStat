param(
  [Parameter(Mandatory = $false)][string]$ManifestPath = "",
  [Parameter(Mandatory = $false)][string]$OutputRoot = "",
  [Parameter(Mandatory = $false)][string]$RunA = "",
  [Parameter(Mandatory = $false)][string]$RunB = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
  $ManifestPath = Join-Path $PSScriptRoot "pack_manifest.json"
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
  $OutputRoot = Join-Path $PSScriptRoot "artifacts"
}

function Test-HasProperty {
  param(
    [Parameter(Mandatory = $true)][object]$Object,
    [Parameter(Mandatory = $true)][string]$Name
  )

  return ($null -ne $Object.PSObject.Properties[$Name])
}

function Get-RequiredString {
  param(
    [Parameter(Mandatory = $true)][object]$Object,
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $false)][string]$Context = "object"
  )

  if (-not (Test-HasProperty -Object $Object -Name $Name)) {
    throw "Missing required field '$Name' in $Context."
  }

  $value = [string]$Object.$Name
  if ([string]::IsNullOrWhiteSpace($value)) {
    throw "Field '$Name' in $Context cannot be empty."
  }

  return $value.Trim()
}

function Find-RepoRoot {
  param(
    [Parameter(Mandatory = $true)][string]$startDir
  )

  $cur = (Resolve-Path -LiteralPath $startDir).Path
  for ($i = 0; $i -lt 8; $i++) {
    if (Test-Path -LiteralPath (Join-Path $cur "AGENTS.md")) {
      return $cur
    }

    $parent = Split-Path -Parent $cur
    if ($parent -eq $cur) {
      break
    }
    $cur = $parent
  }

  throw "Repo root not found (missing AGENTS.md) walking up from $startDir"
}

function Is-CaseEnabled {
  param(
    [Parameter(Mandatory = $true)][object]$Case
  )

  if (-not (Test-HasProperty -Object $Case -Name "enabled")) {
    return $true
  }

  $raw = $Case.enabled
  if ($null -eq $raw) {
    return $true
  }

  if ($raw -is [bool]) {
    return [bool]$raw
  }

  $txt = ([string]$raw).Trim().ToLowerInvariant()
  if ($txt -eq "" -or $txt -eq "true" -or $txt -eq "1" -or $txt -eq "yes") {
    return $true
  }

  if ($txt -eq "false" -or $txt -eq "0" -or $txt -eq "no") {
    return $false
  }

  throw "Invalid boolean value for case.enabled: '$raw'"
}

function Normalize-ArtifactText {
  param(
    [Parameter(Mandatory = $true)][string]$Text
  )

  $normalized = $Text
  $normalized = $normalized -replace "`r`n", "`n"
  $normalized = $normalized -replace '\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]\s*', '[TS] '
  $normalized = $normalized -replace 'RUN_ID=\d{8}_\d{6}', 'RUN_ID=[RUN]'
  $normalized = $normalized -replace 'MACHINE=.*? USER=.*?(\r?\n)', 'MACHINE=[M] USER=[U]$1'
  $normalized = $normalized -replace '[A-Z]:\\[^ \r\n]+', '[PATH]'
  $normalized = $normalized.TrimEnd("`r", "`n")
  return ($normalized + "`n")
}

function Get-StableSha256 {
  param(
    [Parameter(Mandatory = $true)][string]$Text
  )

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hashBytes = $sha.ComputeHash($bytes)
  }
  finally {
    $sha.Dispose()
  }

  return ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ""
}

function Emit-Diff {
  param(
    [Parameter(Mandatory = $true)][string]$LeftPath,
    [Parameter(Mandatory = $true)][string]$RightPath,
    [Parameter(Mandatory = $true)][string]$DiffPath
  )

  $git = Get-Command git -ErrorAction SilentlyContinue
  if ($git) {
    $diffOut = & git diff --no-index -- $LeftPath $RightPath 2>&1
    $gitExit = $LASTEXITCODE
    if ($gitExit -gt 1) {
      throw "git diff failed for '$LeftPath' and '$RightPath' with exit code $gitExit."
    }
    ($diffOut | ForEach-Object { [string]$_ }) | Set-Content -LiteralPath $DiffPath -Encoding ASCII
    return
  }

  $left = Get-Content -LiteralPath $LeftPath
  $right = Get-Content -LiteralPath $RightPath
  $cmp = Compare-Object -ReferenceObject $left -DifferenceObject $right -IncludeEqual:$false
  if ($null -eq $cmp -or $cmp.Count -eq 0) {
    "(No diff output; files reported as mismatched by hash.)" | Set-Content -LiteralPath $DiffPath -Encoding ASCII
    return
  }

  ($cmp | ForEach-Object { [string]$_ }) | Set-Content -LiteralPath $DiffPath -Encoding ASCII
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
  throw "Manifest file not found: $ManifestPath"
}

$manifestFullPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifest = Get-Content -LiteralPath $manifestFullPath -Raw | ConvertFrom-Json
if ($null -eq $manifest) {
  throw "Manifest could not be parsed: $manifestFullPath"
}

$schemaVersion = [int](Get-RequiredString -Object $manifest -Name "schema_version" -Context "manifest")
if ($schemaVersion -ne 1) {
  throw "Unsupported schema_version '$schemaVersion'. Expected 1."
}

$packName = Get-RequiredString -Object $manifest -Name "pack_name" -Context "manifest"

if (-not (Test-HasProperty -Object $manifest -Name "cases")) {
  throw "Manifest is missing 'cases'."
}

$enabledCases = New-Object System.Collections.Generic.List[object]
foreach ($case in @($manifest.cases)) {
  if (Is-CaseEnabled -Case $case) {
    $enabledCases.Add($case) | Out-Null
  }
}
if ($enabledCases.Count -eq 0) {
  throw "No enabled cases found in manifest."
}

$repoRoot = Find-RepoRoot -startDir $PSScriptRoot
$outputRootFull = ""
if ([System.IO.Path]::IsPathRooted($OutputRoot)) {
  $outputRootFull = [System.IO.Path]::GetFullPath($OutputRoot)
}
else {
  $outputRootFull = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputRoot))
}
if (-not (Test-Path -LiteralPath $outputRootFull)) {
  throw "Output root not found: $outputRootFull"
}

if ([string]::IsNullOrWhiteSpace($RunA) -or [string]::IsNullOrWhiteSpace($RunB)) {
  $candidateRuns = Get-ChildItem -LiteralPath $outputRootFull -Directory |
    Where-Object { $_.Name -ne "compare" } |
    Sort-Object LastWriteTimeUtc

  if ($candidateRuns.Count -lt 2) {
    throw "RunA/RunB were not supplied and fewer than 2 run directories exist under $outputRootFull."
  }

  if ([string]::IsNullOrWhiteSpace($RunA)) {
    $RunA = $candidateRuns[$candidateRuns.Count - 2].Name
  }
  if ([string]::IsNullOrWhiteSpace($RunB)) {
    $RunB = $candidateRuns[$candidateRuns.Count - 1].Name
  }
}

$runDirA = Join-Path $outputRootFull $RunA
$runDirB = Join-Path $outputRootFull $RunB
if (-not (Test-Path -LiteralPath $runDirA)) { throw "RunA directory not found: $runDirA" }
if (-not (Test-Path -LiteralPath $runDirB)) { throw "RunB directory not found: $runDirB" }

$caseDirA = Join-Path $runDirA "cases"
$caseDirB = Join-Path $runDirB "cases"
if (-not (Test-Path -LiteralPath $caseDirA)) { throw "RunA cases directory not found: $caseDirA" }
if (-not (Test-Path -LiteralPath $caseDirB)) { throw "RunB cases directory not found: $caseDirB" }

$compareRoot = Join-Path (Join-Path $outputRootFull "compare") ($RunA + "__" + $RunB)
$normDir = Join-Path $compareRoot "normalized"
$diffDir = Join-Path $compareRoot "diffs"
New-Item -ItemType Directory -Force -Path $normDir | Out-Null
New-Item -ItemType Directory -Force -Path $diffDir | Out-Null

$hashByCaseA = @{}
$hashByCaseB = @{}
$normPathA = @{}
$normPathB = @{}
$caseResults = New-Object System.Collections.Generic.List[object]
$mismatchCount = 0
$seenCaseIds = @{}

foreach ($case in $enabledCases) {
  $caseId = Get-RequiredString -Object $case -Name "id" -Context "case"
  if ($seenCaseIds.ContainsKey($caseId)) {
    throw "Duplicate case id found: $caseId"
  }
  $seenCaseIds[$caseId] = $true

  $artifactRel = Get-RequiredString -Object $case -Name "artifact" -Context "case '$caseId'"
  if ([System.IO.Path]::IsPathRooted($artifactRel)) {
    throw "Case '$caseId' has rooted artifact path; only relative paths are allowed."
  }

  $artifactA = Join-Path $caseDirA $artifactRel
  $artifactB = Join-Path $caseDirB $artifactRel
  if (-not (Test-Path -LiteralPath $artifactA)) { throw "Missing RunA artifact for case '$caseId': $artifactA" }
  if (-not (Test-Path -LiteralPath $artifactB)) { throw "Missing RunB artifact for case '$caseId': $artifactB" }

  $rawA = Get-Content -LiteralPath $artifactA -Raw
  $rawB = Get-Content -LiteralPath $artifactB -Raw

  $normA = Normalize-ArtifactText -Text $rawA
  $normB = Normalize-ArtifactText -Text $rawB

  $normAPath = Join-Path $normDir ($caseId + ".runA.norm.txt")
  $normBPath = Join-Path $normDir ($caseId + ".runB.norm.txt")
  $normA | Set-Content -LiteralPath $normAPath -Encoding ASCII
  $normB | Set-Content -LiteralPath $normBPath -Encoding ASCII

  $hashA = Get-StableSha256 -Text $normA
  $hashB = Get-StableSha256 -Text $normB
  $equal = ($hashA -eq $hashB)

  $hashByCaseA[$caseId] = $hashA
  $hashByCaseB[$caseId] = $hashB
  $normPathA[$caseId] = $normAPath
  $normPathB[$caseId] = $normBPath

  $diffFile = ""
  if (-not $equal) {
    $mismatchCount++
    $diffFile = Join-Path $diffDir ("case_" + $caseId + ".diff.txt")
    Emit-Diff -LeftPath $normAPath -RightPath $normBPath -DiffPath $diffFile
  }

  $caseResults.Add([pscustomobject]@{
      case_id = $caseId
      artifact = $artifactRel
      runA_sha256 = $hashA
      runB_sha256 = $hashB
      equal = $equal
      diff_file = $diffFile
    }) | Out-Null
}

$withinResults = New-Object System.Collections.Generic.List[object]
if ((Test-HasProperty -Object $manifest -Name "within_run_equal") -and $null -ne $manifest.within_run_equal) {
  foreach ($check in @($manifest.within_run_equal)) {
    $checkId = Get-RequiredString -Object $check -Name "id" -Context "within_run_equal check"
    $leftCase = Get-RequiredString -Object $check -Name "left_case" -Context "within_run_equal '$checkId'"
    $rightCase = Get-RequiredString -Object $check -Name "right_case" -Context "within_run_equal '$checkId'"

    if (-not $hashByCaseA.ContainsKey($leftCase)) {
      throw "within_run_equal '$checkId' references unknown left_case '$leftCase'."
    }
    if (-not $hashByCaseA.ContainsKey($rightCase)) {
      throw "within_run_equal '$checkId' references unknown right_case '$rightCase'."
    }

    $runAEqual = ($hashByCaseA[$leftCase] -eq $hashByCaseA[$rightCase])
    $runBEqual = ($hashByCaseB[$leftCase] -eq $hashByCaseB[$rightCase])
    $runADiff = ""
    $runBDiff = ""

    if (-not $runAEqual) {
      $mismatchCount++
      $runADiff = Join-Path $diffDir ("within_" + $checkId + "_runA.diff.txt")
      Emit-Diff -LeftPath $normPathA[$leftCase] -RightPath $normPathA[$rightCase] -DiffPath $runADiff
    }

    if (-not $runBEqual) {
      $mismatchCount++
      $runBDiff = Join-Path $diffDir ("within_" + $checkId + "_runB.diff.txt")
      Emit-Diff -LeftPath $normPathB[$leftCase] -RightPath $normPathB[$rightCase] -DiffPath $runBDiff
    }

    $withinResults.Add([pscustomobject]@{
        check_id = $checkId
        left_case = $leftCase
        right_case = $rightCase
        runA_equal = $runAEqual
        runB_equal = $runBEqual
        runA_diff_file = $runADiff
        runB_diff_file = $runBDiff
      }) | Out-Null
  }
}

$overallPass = ($mismatchCount -eq 0)
$reportTxtPath = Join-Path $compareRoot "compare_report.txt"
$reportJsonPath = Join-Path $compareRoot "compare_report.json"

$report = [ordered]@{
  schema_version = 1
  pack_name = $packName
  manifest_path = $manifestFullPath
  run_a = $RunA
  run_b = $RunB
  mismatch_count = $mismatchCount
  pass = $overallPass
  case_results = $caseResults.ToArray()
  within_run_results = $withinResults.ToArray()
}

$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportJsonPath -Encoding ASCII

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("PACK_NAME=$packName") | Out-Null
$lines.Add("RUN_A=$RunA") | Out-Null
$lines.Add("RUN_B=$RunB") | Out-Null
$lines.Add("MISMATCH_COUNT=$mismatchCount") | Out-Null
$lines.Add("PACK_PASS=$overallPass") | Out-Null
foreach ($r in $caseResults) {
  $lines.Add("CASE=$($r.case_id) RUN_A_SHA256=$($r.runA_sha256) RUN_B_SHA256=$($r.runB_sha256) EQUAL=$($r.equal) DIFF=$($r.diff_file)") | Out-Null
}
foreach ($w in $withinResults) {
  $lines.Add("WITHIN=$($w.check_id) LEFT=$($w.left_case) RIGHT=$($w.right_case) RUN_A_EQUAL=$($w.runA_equal) RUN_B_EQUAL=$($w.runB_equal) RUN_A_DIFF=$($w.runA_diff_file) RUN_B_DIFF=$($w.runB_diff_file)") | Out-Null
}
$lines | Set-Content -LiteralPath $reportTxtPath -Encoding ASCII

Write-Host "COMPARE_ROOT=$compareRoot"
Write-Host "REPORT_TXT=$reportTxtPath"
Write-Host "REPORT_JSON=$reportJsonPath"
Write-Host "PACK_PASS=$overallPass"

if (-not $overallPass) {
  exit 2
}

exit 0
