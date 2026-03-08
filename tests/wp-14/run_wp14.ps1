param(
  [Parameter(Mandatory = $false)][string]$RunLabel = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Find-RepoRoot {
  param(
    [Parameter(Mandatory = $true)][string]$StartDir
  )

  $current = (Resolve-Path -LiteralPath $StartDir).Path
  while ($true) {
    if (Test-Path -LiteralPath (Join-Path $current "AGENTS.md")) {
      return $current
    }

    $parent = Split-Path -Parent $current
    if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
      break
    }

    $current = $parent
  }

  throw "Repo root not found (missing AGENTS.md) walking up from $StartDir"
}

function New-RunLabel {
  param(
    [Parameter(Mandatory = $false)][string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return "run_" + [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")
  }

  return ($Value.Trim() -replace '[^A-Za-z0-9._-]', '_')
}

function Get-SafeToken {
  param(
    [Parameter(Mandatory = $true)][string]$Value
  )

  $clean = ($Value.Trim() -replace '[^A-Za-z0-9._-]', '_')
  if ([string]::IsNullOrWhiteSpace($clean)) {
    return "case"
  }
  return $clean
}

function Read-KeyValueFromOutput {
  param(
    [Parameter(Mandatory = $true)][string[]]$Lines,
    [Parameter(Mandatory = $true)][string]$Key
  )

  foreach ($line in $Lines) {
    if ($line -match ("^{0}=(.+)$" -f [System.Text.RegularExpressions.Regex]::Escape($Key))) {
      return $Matches[1].Trim()
    }
  }

  return ""
}

function Invoke-ValidationCase {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$Group,
    [Parameter(Mandatory = $true)][string]$ValidatorPath,
    [Parameter(Mandatory = $true)][string]$TargetPath,
    [Parameter(Mandatory = $true)][bool]$ExpectedPass,
    [Parameter(Mandatory = $true)][string]$RunLabel,
    [Parameter(Mandatory = $true)][string]$LogsDir
  )

  $caseLabel = Get-SafeToken -Value $Name
  $logPath = Join-Path $LogsDir ("{0}.log" -f $caseLabel)

  $stdout = & pwsh -NoProfile -ExecutionPolicy Bypass -File $ValidatorPath -Path $TargetPath -RunLabel $RunLabel -CaseLabel $caseLabel 2>&1
  $exitCode = $LASTEXITCODE
  $outputLines = @($stdout | ForEach-Object { [string]$_ })
  if ($outputLines.Count -eq 0) {
    $outputLines = @("")
  }
  $outputLines | Set-Content -LiteralPath $logPath -Encoding ASCII

  $actualPass = ($exitCode -eq 0)
  $expectationMatched = ($actualPass -eq $ExpectedPass)
  $jsonSummary = Read-KeyValueFromOutput -Lines $outputLines -Key "JSON_SUMMARY"
  $txtSummary = Read-KeyValueFromOutput -Lines $outputLines -Key "TXT_SUMMARY"
  $resultWord = Read-KeyValueFromOutput -Lines $outputLines -Key "RESULT"

  return [pscustomobject]@{
    name = $Name
    group = $Group
    validator = $ValidatorPath
    target_path = $TargetPath
    expected_pass = $ExpectedPass
    actual_pass = $actualPass
    expectation_matched = $expectationMatched
    exit_code = $exitCode
    validator_result = $resultWord
    json_summary = $jsonSummary
    txt_summary = $txtSummary
    log_file = $logPath
  }
}

$repoRoot = Find-RepoRoot -StartDir $PSScriptRoot
$runLabelFinal = New-RunLabel -Value $RunLabel

$validatorsDir = Join-Path $PSScriptRoot "contract-validators"
$templateValidator = Join-Path $validatorsDir "validate_template_config_contract.ps1"
$mappingsValidator = Join-Path $validatorsDir "validate_mappings_contract.ps1"
$overridesValidator = Join-Path $validatorsDir "validate_static_overrides_contract.ps1"

foreach ($validator in @($templateValidator, $mappingsValidator, $overridesValidator)) {
  if (-not (Test-Path -LiteralPath $validator)) {
    throw "Validator missing: $validator"
  }
}

$artifactRoot = Join-Path $validatorsDir "artifacts"
$runDir = Join-Path $artifactRoot $runLabelFinal
$logsDir = Join-Path $runDir "logs"
New-Item -ItemType Directory -Force -Path $logsDir | Out-Null

$fixturesRoot = Join-Path $PSScriptRoot "fixtures"
$goodFixturesDir = Join-Path $fixturesRoot "good"
$badFixturesDir = Join-Path $fixturesRoot "bad"

if (-not (Test-Path -LiteralPath $goodFixturesDir)) { throw "Missing good fixtures directory: $goodFixturesDir" }
if (-not (Test-Path -LiteralPath $badFixturesDir)) { throw "Missing bad fixtures directory: $badFixturesDir" }

$goodFixtures = @(Get-ChildItem -LiteralPath $goodFixturesDir -File -Filter *.ini | Sort-Object Name)
$badFixtures = @(Get-ChildItem -LiteralPath $badFixturesDir -File -Filter *.ini | Sort-Object Name)

if ($goodFixtures.Count -lt 1) { throw "Expected at least 1 good fixture. Found: $($goodFixtures.Count)" }
if ($badFixtures.Count -lt 2) { throw "Expected at least 2 bad fixtures. Found: $($badFixtures.Count)" }

$startedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
$results = [System.Collections.Generic.List[object]]::new()

$repoCases = @(
  @{
    Name = "repo_template_config"
    Group = "repo"
    Validator = $templateValidator
    Target = (Join-Path $repoRoot "SmartStat_TemplateConfig.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_mlb"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_Mappings.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_nba"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_MappingsNBA.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_nhl"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_MappingsNHL.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_learn_mlb"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_Mappings.learn.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_learn_nba"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_MappingsNBA.learn.ini")
    Expected = $true
  },
  @{
    Name = "repo_mappings_learn_nhl"
    Group = "repo"
    Validator = $mappingsValidator
    Target = (Join-Path $repoRoot "SmartStat_MappingsNHL.learn.ini")
    Expected = $true
  },
  @{
    Name = "repo_static_overrides"
    Group = "repo"
    Validator = $overridesValidator
    Target = (Join-Path $repoRoot "SmartStat_StaticOverrides.ini")
    Expected = $true
  }
)

foreach ($case in $repoCases) {
  $result = Invoke-ValidationCase `
    -Name $case.Name `
    -Group $case.Group `
    -ValidatorPath $case.Validator `
    -TargetPath $case.Target `
    -ExpectedPass $case.Expected `
    -RunLabel $runLabelFinal `
    -LogsDir $logsDir
  $results.Add($result) | Out-Null
}

foreach ($fixture in $goodFixtures) {
  $result = Invoke-ValidationCase `
    -Name ("fixture_good_" + [System.IO.Path]::GetFileNameWithoutExtension($fixture.Name)) `
    -Group "fixture_good" `
    -ValidatorPath $templateValidator `
    -TargetPath $fixture.FullName `
    -ExpectedPass $true `
    -RunLabel $runLabelFinal `
    -LogsDir $logsDir
  $results.Add($result) | Out-Null
}

foreach ($fixture in $badFixtures) {
  $result = Invoke-ValidationCase `
    -Name ("fixture_bad_" + [System.IO.Path]::GetFileNameWithoutExtension($fixture.Name)) `
    -Group "fixture_bad" `
    -ValidatorPath $templateValidator `
    -TargetPath $fixture.FullName `
    -ExpectedPass $false `
    -RunLabel $runLabelFinal `
    -LogsDir $logsDir
  $results.Add($result) | Out-Null
}

$expectationFailures = 0
$repoGateFailures = 0
foreach ($result in $results) {
  if (-not $result.expectation_matched) {
    $expectationFailures++
  }
  if ($result.group -eq "repo" -and -not $result.actual_pass) {
    $repoGateFailures++
  }
}

$endedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
$runPass = ($expectationFailures -eq 0 -and $repoGateFailures -eq 0)

$summaryJson = Join-Path $runDir "run_summary.json"
$summaryTxt = Join-Path $runDir "run_summary.txt"

$summary = [ordered]@{
  schema_version = 1
  run_label = $runLabelFinal
  repo_root = $repoRoot
  started_utc = $startedUtc
  ended_utc = $endedUtc
  repo_case_count = $repoCases.Count
  good_fixture_count = $goodFixtures.Count
  bad_fixture_count = $badFixtures.Count
  case_count = $results.Count
  expectation_failures = $expectationFailures
  repo_gate_failures = $repoGateFailures
  run_pass = $runPass
  results = $results.ToArray()
}

$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJson -Encoding ASCII

$txtLines = [System.Collections.Generic.List[string]]::new()
$txtLines.Add("RUN_LABEL=$runLabelFinal") | Out-Null
$txtLines.Add("RUN_DIR=$runDir") | Out-Null
$txtLines.Add("REPO_CASE_COUNT=$($repoCases.Count)") | Out-Null
$txtLines.Add("GOOD_FIXTURE_COUNT=$($goodFixtures.Count)") | Out-Null
$txtLines.Add("BAD_FIXTURE_COUNT=$($badFixtures.Count)") | Out-Null
$txtLines.Add("CASE_COUNT=$($results.Count)") | Out-Null
$txtLines.Add("EXPECTATION_FAILURES=$expectationFailures") | Out-Null
$txtLines.Add("REPO_GATE_FAILURES=$repoGateFailures") | Out-Null
$txtLines.Add("RUN_PASS=$runPass") | Out-Null
foreach ($result in $results) {
  $txtLines.Add("CASE=$($result.name) GROUP=$($result.group) EXPECTED_PASS=$($result.expected_pass) ACTUAL_PASS=$($result.actual_pass) MATCHED=$($result.expectation_matched) EXIT_CODE=$($result.exit_code) TARGET=$($result.target_path) LOG=$($result.log_file)") | Out-Null
}
$txtLines | Set-Content -LiteralPath $summaryTxt -Encoding ASCII

Write-Host "RUN_LABEL=$runLabelFinal"
Write-Host "RUN_DIR=$runDir"
Write-Host "SUMMARY_JSON=$summaryJson"
Write-Host "SUMMARY_TXT=$summaryTxt"
Write-Host "CASE_COUNT=$($results.Count)"
Write-Host "EXPECTATION_FAILURES=$expectationFailures"
Write-Host "REPO_GATE_FAILURES=$repoGateFailures"
Write-Host "RUN_PASS=$runPass"

if ($runPass) {
  exit 0
}

exit 2
