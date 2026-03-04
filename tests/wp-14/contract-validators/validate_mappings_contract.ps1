param(
  [Parameter(Mandatory = $false)][string]$Path = "",
  [Parameter(Mandatory = $false)][string]$RunLabel = "",
  [Parameter(Mandatory = $false)][string]$CaseLabel = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$commonPath = Join-Path $PSScriptRoot "common_contract_validator.ps1"
if (-not (Test-Path -LiteralPath $commonPath)) {
  throw "Missing helper script: $commonPath"
}
. $commonPath

$validatorId = "validate_mappings_contract"
$repoRoot = Find-RepoRoot -StartDir $PSScriptRoot
$resolvedPath = Resolve-InputPath -RepoRoot $repoRoot -Path $Path -DefaultFileName "SmartStat_Mappings.ini"
$runLabelFinal = New-RunLabel -RunLabel $RunLabel
$caseTokenSource = if ([string]::IsNullOrWhiteSpace($CaseLabel)) { [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) } else { $CaseLabel }
$caseToken = Get-SafeToken -Value $caseTokenSource -Fallback "mappings"

$artifactDir = Join-Path (Join-Path $PSScriptRoot "artifacts") $runLabelFinal
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null
$jsonPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.json"
$txtPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.txt"

$issues = New-IssueList
$parseResult = Parse-IniDocument -Path $resolvedPath
foreach ($issue in $parseResult.issues) {
  $issues.Add($issue) | Out-Null
}

$duplicateKeyCount = 0

foreach ($section in $parseResult.sections) {
  $keysInSection = @{}
  foreach ($entry in $section.keys) {
    if (-not $keysInSection.ContainsKey($entry.key_normalized)) {
      $keysInSection[$entry.key_normalized] = [System.Collections.Generic.List[object]]::new()
    }
    $keysInSection[$entry.key_normalized].Add($entry) | Out-Null
  }

  foreach ($keyGroup in $keysInSection.GetEnumerator()) {
    if ($keyGroup.Value.Count -gt 1) {
      for ($dupIdx = 1; $dupIdx -lt $keyGroup.Value.Count; $dupIdx++) {
        $dup = $keyGroup.Value[$dupIdx]
        $duplicateKeyCount++
        Add-Issue -Issues $issues -Code "DUPLICATE_KEY" -Line $dup.line -Section $section.name -Key $dup.key -Message "Duplicate key '$($dup.key)' in section '$($section.name)'."
      }
    }
  }
}

$validationPass = ($issues.Count -eq 0)
$result = [ordered]@{
  schema_version = 1
  validator = $validatorId
  run_label = $runLabelFinal
  case_label = $CaseLabel
  repo_root = $repoRoot
  target_path = $parseResult.path
  exists = $parseResult.exists
  parseable = $parseResult.parseable
  utf8_bom = $parseResult.utf8_bom
  line_ending = $parseResult.line_ending
  section_count = $parseResult.section_count
  key_count = $parseResult.key_count
  duplicate_key_count = $duplicateKeyCount
  issue_count = $issues.Count
  valid = $validationPass
  issues = $issues.ToArray()
}

$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding ASCII

$txtLines = [System.Collections.Generic.List[string]]::new()
$txtLines.Add("VALIDATOR=$validatorId") | Out-Null
$txtLines.Add("RUN_LABEL=$runLabelFinal") | Out-Null
$txtLines.Add("CASE_LABEL=$CaseLabel") | Out-Null
$txtLines.Add("TARGET_PATH=$($parseResult.path)") | Out-Null
$txtLines.Add("PARSEABLE=$($parseResult.parseable)") | Out-Null
$txtLines.Add("SECTION_COUNT=$($parseResult.section_count)") | Out-Null
$txtLines.Add("KEY_COUNT=$($parseResult.key_count)") | Out-Null
$txtLines.Add("DUPLICATE_KEY_COUNT=$duplicateKeyCount") | Out-Null
$txtLines.Add("ISSUE_COUNT=$($issues.Count)") | Out-Null
$txtLines.Add("RESULT=$(if ($validationPass) { 'PASS' } else { 'FAIL' })") | Out-Null
foreach ($issue in $issues) {
  $txtLines.Add("ISSUE LINE=$($issue.line) CODE=$($issue.code) SECTION=$($issue.section) KEY=$($issue.key) MSG=$($issue.message)") | Out-Null
}
$txtLines | Set-Content -LiteralPath $txtPath -Encoding ASCII

Write-Host "VALIDATOR=$validatorId"
Write-Host "RUN_LABEL=$runLabelFinal"
Write-Host "CASE_LABEL=$CaseLabel"
Write-Host "TARGET_PATH=$($parseResult.path)"
Write-Host "PARSEABLE=$($parseResult.parseable)"
Write-Host "SECTION_COUNT=$($parseResult.section_count)"
Write-Host "KEY_COUNT=$($parseResult.key_count)"
Write-Host "DUPLICATE_KEY_COUNT=$duplicateKeyCount"
Write-Host "ISSUE_COUNT=$($issues.Count)"
Write-Host ("RESULT={0}" -f ($(if ($validationPass) { "PASS" } else { "FAIL" })))
Write-Host "JSON_SUMMARY=$jsonPath"
Write-Host "TXT_SUMMARY=$txtPath"

if ($validationPass) {
  exit 0
}

exit 2
