param(
  [Parameter(Mandatory = $false)][string]$RunLabel = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Find-RepoRoot {
  param(
    [Parameter(Mandatory = $true)][string]$StartDir
  )

  $cur = (Resolve-Path -LiteralPath $StartDir).Path
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

  throw "Repo root not found (missing AGENTS.md) walking up from $StartDir"
}

function Invoke-ValidationCase {
  param(
    [Parameter(Mandatory = $true)][string]$ValidatorPath,
    [Parameter(Mandatory = $true)][string]$IniPath,
    [Parameter(Mandatory = $false)][object]$ExpectedPass = $null,
    [Parameter(Mandatory = $true)][string]$CaseGroup,
    [Parameter(Mandatory = $true)][int]$CaseIndex,
    [Parameter(Mandatory = $true)][string]$CaseOutputDir
  )

  $leafName = [System.IO.Path]::GetFileName($IniPath)
  $safeLeaf = ($leafName -replace '[^A-Za-z0-9._-]', '_')
  $caseFile = ("{0:D2}_{1}_{2}.txt" -f $CaseIndex, $CaseGroup, $safeLeaf)
  $outputPath = Join-Path $CaseOutputDir $caseFile

  $stdout = & powershell -NoProfile -ExecutionPolicy Bypass -File $ValidatorPath -IniPath $IniPath 2>&1
  $exitCode = $LASTEXITCODE

  if ($null -eq $stdout) {
    "" | Set-Content -LiteralPath $outputPath -Encoding ASCII
  }
  else {
    ($stdout | ForEach-Object { [string]$_ }) | Set-Content -LiteralPath $outputPath -Encoding ASCII
  }

  $actualPass = ($exitCode -eq 0)
  $expectationMatched = $true
  if ($null -ne $ExpectedPass) {
    $expectationMatched = ($actualPass -eq [bool]$ExpectedPass)
  }

  return [pscustomobject]@{
    case_index = $CaseIndex
    group = $CaseGroup
    ini_path = $IniPath
    expected_pass = $ExpectedPass
    actual_pass = $actualPass
    expectation_matched = $expectationMatched
    exit_code = $exitCode
    output_file = $outputPath
  }
}

$repoRoot = Find-RepoRoot -StartDir $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($RunLabel)) {
  $RunLabel = "run_" + [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")
}

$validatorPath = Join-Path $PSScriptRoot "validate_learn_ini.ps1"
if (-not (Test-Path -LiteralPath $validatorPath)) {
  throw "Validator script missing: $validatorPath"
}

$goodDir = Join-Path $PSScriptRoot "fixtures\good"
$badDir = Join-Path $PSScriptRoot "fixtures\bad"

if (-not (Test-Path -LiteralPath $goodDir)) { throw "Missing fixture directory: $goodDir" }
if (-not (Test-Path -LiteralPath $badDir)) { throw "Missing fixture directory: $badDir" }

$goodFixtures = @(Get-ChildItem -LiteralPath $goodDir -File -Filter *.ini | Sort-Object Name)
$badFixtures = @(Get-ChildItem -LiteralPath $badDir -File -Filter *.ini | Sort-Object Name)

if ($goodFixtures.Count -lt 2) { throw "Expected at least 2 good fixtures. Found: $($goodFixtures.Count)" }
if ($badFixtures.Count -lt 3) { throw "Expected at least 3 bad fixtures. Found: $($badFixtures.Count)" }

$repoLearn = @()
$repoLearn += @(Get-ChildItem -LiteralPath $repoRoot -File -Filter *Mappings*.learn.ini -ErrorAction SilentlyContinue)
$repoLearn += @(Get-ChildItem -LiteralPath $repoRoot -File -Filter *.learn.ini -ErrorAction SilentlyContinue)
$repoLearn = @($repoLearn | Sort-Object FullName -Unique)

$artifactRoot = Join-Path $PSScriptRoot "artifacts"
$runDir = Join-Path $artifactRoot $RunLabel
$caseOutDir = Join-Path $runDir "cases"
New-Item -ItemType Directory -Force -Path $caseOutDir | Out-Null

$startedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
$results = [System.Collections.Generic.List[object]]::new()
$caseIndex = 1

foreach ($f in $goodFixtures) {
  $result = Invoke-ValidationCase -ValidatorPath $validatorPath -IniPath $f.FullName -ExpectedPass:$true -CaseGroup "good_fixture" -CaseIndex $caseIndex -CaseOutputDir $caseOutDir
  $results.Add($result) | Out-Null
  $caseIndex++
}

foreach ($f in $badFixtures) {
  $result = Invoke-ValidationCase -ValidatorPath $validatorPath -IniPath $f.FullName -ExpectedPass:$false -CaseGroup "bad_fixture" -CaseIndex $caseIndex -CaseOutputDir $caseOutDir
  $results.Add($result) | Out-Null
  $caseIndex++
}

foreach ($f in $repoLearn) {
  # RC-safe: repo learn files are validated read-only and reported as observational (non-gating).
  $result = Invoke-ValidationCase -ValidatorPath $validatorPath -IniPath $f.FullName -CaseGroup "repo_file" -CaseIndex $caseIndex -CaseOutputDir $caseOutDir
  $results.Add($result) | Out-Null
  $caseIndex++
}

$expectationFailures = 0
foreach ($r in $results) {
  if (-not $r.expectation_matched) {
    $expectationFailures++
  }
}

$endedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
$runPass = ($expectationFailures -eq 0)

$summaryJson = Join-Path $runDir "run_summary.json"
$summaryTxt = Join-Path $runDir "run_summary.txt"

$summary = [ordered]@{
  schema_version = 1
  run_label = $RunLabel
  repo_root = $repoRoot
  started_utc = $startedUtc
  ended_utc = $endedUtc
  good_fixture_count = $goodFixtures.Count
  bad_fixture_count = $badFixtures.Count
  repo_learn_file_count = $repoLearn.Count
  case_count = $results.Count
  expectation_failures = $expectationFailures
  run_pass = $runPass
  results = $results.ToArray()
}

$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJson -Encoding ASCII

$txt = [System.Collections.Generic.List[string]]::new()
$txt.Add("RUN_LABEL=$RunLabel") | Out-Null
$txt.Add("RUN_DIR=$runDir") | Out-Null
$txt.Add("GOOD_FIXTURE_COUNT=$($goodFixtures.Count)") | Out-Null
$txt.Add("BAD_FIXTURE_COUNT=$($badFixtures.Count)") | Out-Null
$txt.Add("REPO_LEARN_FILE_COUNT=$($repoLearn.Count)") | Out-Null
$txt.Add("CASE_COUNT=$($results.Count)") | Out-Null
$txt.Add("EXPECTATION_FAILURES=$expectationFailures") | Out-Null
$txt.Add("RUN_PASS=$runPass") | Out-Null
foreach ($r in $results) {
  $txt.Add("CASE=$($r.case_index) GROUP=$($r.group) EXPECTED_PASS=$($r.expected_pass) ACTUAL_PASS=$($r.actual_pass) MATCHED=$($r.expectation_matched) EXIT_CODE=$($r.exit_code) PATH=$($r.ini_path)") | Out-Null
}
$txt | Set-Content -LiteralPath $summaryTxt -Encoding ASCII

Write-Host "RUN_LABEL=$RunLabel"
Write-Host "RUN_DIR=$runDir"
Write-Host "SUMMARY_JSON=$summaryJson"
Write-Host "SUMMARY_TXT=$summaryTxt"
Write-Host "RUN_PASS=$runPass"

if (-not $runPass) {
  exit 2
}

exit 0
