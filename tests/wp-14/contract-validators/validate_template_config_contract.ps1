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

function New-KeyEntryMap {
  return @{}
}

function Add-KeyEntry {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Map,
    [Parameter(Mandatory = $true)][object]$Entry
  )

  $key = [string]$Entry.key_normalized
  if (-not $Map.ContainsKey($key)) {
    $Map[$key] = [System.Collections.Generic.List[object]]::new()
  }
  $Map[$key].Add($Entry) | Out-Null
}

function Test-TabfieldToken {
  param(
    [Parameter(Mandatory = $true)][string]$Token
  )

  return [System.Text.RegularExpressions.Regex]::IsMatch($Token, "^[A-Z][0-9]{4}(?:-[A-Za-z0-9_]+)?$")
}

function Test-TabfieldListValue {
  param(
    [Parameter(Mandatory = $true)][string]$Value
  )

  if ($Value -eq "none") {
    return $true
  }

  $parts = $Value.Split(",")
  if ($parts.Count -eq 0) {
    return $false
  }

  foreach ($part in $parts) {
    $token = $part.Trim()
    if ([string]::IsNullOrWhiteSpace($token)) {
      return $false
    }

    if (-not (Test-TabfieldToken -Token $token)) {
      return $false
    }
  }

  return $true
}

function Test-OutputMapValue {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return $true
  }

  $parts = $Value.Split(",")
  foreach ($part in $parts) {
    $token = $part.Trim()
    if ([string]::IsNullOrWhiteSpace($token)) {
      return $false
    }

    if (-not [System.Text.RegularExpressions.Regex]::IsMatch($token, "^[A-Z][0-9]{4}:[0-9]+:[0-9]+$")) {
      return $false
    }
  }

  return $true
}

$validatorId = "validate_template_config_contract"
$contractVersion = "1.0.0"
$requiredKeys = @("config_id", "qualifier", "filter_tabfields", "category_tabfields", "row_limit", "output_map")
$allowedEmptyKeys = @("output_map")
$allowedNonTemplateSections = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$allowedNonTemplateSections.Add("_BOM_GUARD_") | Out-Null
$allowedNonTemplateSections.Add("SECURITY") | Out-Null

$repoRoot = Find-RepoRoot -StartDir $PSScriptRoot
$resolvedPath = Resolve-InputPath -RepoRoot $repoRoot -Path $Path -DefaultFileName "SmartStat_TemplateConfig.ini"
$runLabelFinal = New-RunLabel -RunLabel $RunLabel
$caseTokenSource = if ([string]::IsNullOrWhiteSpace($CaseLabel)) { [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) } else { $CaseLabel }
$caseToken = Get-SafeToken -Value $caseTokenSource -Fallback "template_config"

$artifactDir = Join-Path (Join-Path $PSScriptRoot "artifacts") $runLabelFinal
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null
$jsonPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.json"
$txtPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.txt"

$issues = New-IssueList
$parseResult = Parse-IniDocument -Path $resolvedPath
foreach ($issue in $parseResult.issues) {
  $issues.Add($issue) | Out-Null
}

$templateSectionCount = 0
$duplicateTemplateConfigIdCount = 0
$templateConfigIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($section in $parseResult.sections) {
  if ($section.name.StartsWith("TEMPLATE:", [System.StringComparison]::OrdinalIgnoreCase)) {
    $templateSectionCount++

    if (-not [System.Text.RegularExpressions.Regex]::IsMatch($section.name, "^TEMPLATE:[A-Za-z0-9_]+$")) {
      Add-Issue -Issues $issues -Code "INVALID_TEMPLATE_HEADER" -Line $section.header_line -Section $section.name -Key "" -Message "Template section name must match TEMPLATE:<name> with [A-Za-z0-9_]+."
    }

    $keyMap = New-KeyEntryMap
    $encounteredOrder = [System.Collections.Generic.List[string]]::new()
    $encounteredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($entry in $section.keys) {
      Add-KeyEntry -Map $keyMap -Entry $entry

      if (-not $encounteredSet.Contains($entry.key_normalized)) {
        $encounteredSet.Add($entry.key_normalized) | Out-Null
        $encounteredOrder.Add($entry.key_normalized) | Out-Null
      }

      if (-not ($requiredKeys -contains $entry.key_normalized)) {
        Add-Issue -Issues $issues -Code "UNKNOWN_TEMPLATE_KEY" -Line $entry.line -Section $section.name -Key $entry.key -Message "Unknown key '$($entry.key)' in template section; contract 1.0.0 allows only required keys."
      }
    }

    foreach ($mapEntry in $keyMap.GetEnumerator()) {
      if ($mapEntry.Value.Count -gt 1) {
        for ($dupIdx = 1; $dupIdx -lt $mapEntry.Value.Count; $dupIdx++) {
          $dupEntry = $mapEntry.Value[$dupIdx]
          Add-Issue -Issues $issues -Code "DUPLICATE_TEMPLATE_KEY" -Line $dupEntry.line -Section $section.name -Key $dupEntry.key -Message "Duplicate key '$($dupEntry.key)' in template section."
        }
      }
    }

    $hasAllRequired = $true
    foreach ($requiredKey in $requiredKeys) {
      if (-not $keyMap.ContainsKey($requiredKey)) {
        Add-Issue -Issues $issues -Code "MISSING_REQUIRED_KEY" -Line $section.header_line -Section $section.name -Key $requiredKey -Message "Missing required key '$requiredKey'."
        $hasAllRequired = $false
      }
      else {
        $entryList = $keyMap[$requiredKey]
        foreach ($entry in $entryList) {
          if ([string]::IsNullOrWhiteSpace($entry.value) -and -not ($allowedEmptyKeys -contains $requiredKey)) {
            Add-Issue -Issues $issues -Code "EMPTY_VALUE_NOT_ALLOWED" -Line $entry.line -Section $section.name -Key $entry.key -Message "Key '$($entry.key)' cannot be empty."
          }
        }
      }
    }

    if ($hasAllRequired) {
      $requiredEncountered = @()
      foreach ($keyName in $encounteredOrder) {
        if ($requiredKeys -contains $keyName) {
          $requiredEncountered += $keyName
        }
      }

      $expectedOrder = [string]::Join(",", $requiredKeys)
      $actualOrder = [string]::Join(",", $requiredEncountered)
      if ($actualOrder -ne $expectedOrder) {
        Add-Issue -Issues $issues -Code "REQUIRED_KEY_ORDER_VIOLATION" -Line $section.header_line -Section $section.name -Key "" -Message "Required key order mismatch. Expected '$expectedOrder', actual '$actualOrder'."
      }
    }

    if ($keyMap.ContainsKey("config_id")) {
      $configValue = [string]$keyMap["config_id"][0].value
      if (-not [string]::IsNullOrWhiteSpace($configValue) -and -not [System.Text.RegularExpressions.Regex]::IsMatch($configValue, "^[A-Za-z0-9_.-]+$")) {
        Add-Issue -Issues $issues -Code "INVALID_CONFIG_ID_VALUE" -Line $keyMap["config_id"][0].line -Section $section.name -Key "config_id" -Message "config_id must match [A-Za-z0-9_.-]+."
      }
    }

    if ($keyMap.ContainsKey("qualifier")) {
      $qualifierValue = [string]$keyMap["qualifier"][0].value
      if (-not [string]::IsNullOrWhiteSpace($qualifierValue) -and -not [System.Text.RegularExpressions.Regex]::IsMatch($qualifierValue, "^(none|[A-Z][0-9]{4})$")) {
        Add-Issue -Issues $issues -Code "INVALID_QUALIFIER_VALUE" -Line $keyMap["qualifier"][0].line -Section $section.name -Key "qualifier" -Message "qualifier must be 'none' or a tabfield token like B0200."
      }
    }

    if ($keyMap.ContainsKey("filter_tabfields")) {
      $filterValue = [string]$keyMap["filter_tabfields"][0].value
      if (-not [string]::IsNullOrWhiteSpace($filterValue) -and -not (Test-TabfieldListValue -Value $filterValue)) {
        Add-Issue -Issues $issues -Code "INVALID_FILTER_TABFIELDS" -Line $keyMap["filter_tabfields"][0].line -Section $section.name -Key "filter_tabfields" -Message "filter_tabfields must be 'none' or comma-separated tabfield tokens."
      }
    }

    if ($keyMap.ContainsKey("category_tabfields")) {
      $categoryValue = [string]$keyMap["category_tabfields"][0].value
      if (-not [string]::IsNullOrWhiteSpace($categoryValue) -and -not (Test-TabfieldListValue -Value $categoryValue)) {
        Add-Issue -Issues $issues -Code "INVALID_CATEGORY_TABFIELDS" -Line $keyMap["category_tabfields"][0].line -Section $section.name -Key "category_tabfields" -Message "category_tabfields must be comma-separated tabfield tokens."
      }
    }

    if ($keyMap.ContainsKey("row_limit")) {
      $rowValue = [string]$keyMap["row_limit"][0].value
      if (-not [string]::IsNullOrWhiteSpace($rowValue) -and -not [System.Text.RegularExpressions.Regex]::IsMatch($rowValue, "^(none|[A-Z][0-9]{4}(?:-[A-Za-z0-9_]+)?),[1-9][0-9]*$")) {
        Add-Issue -Issues $issues -Code "INVALID_ROW_LIMIT" -Line $keyMap["row_limit"][0].line -Section $section.name -Key "row_limit" -Message "row_limit must match '<token>,<positive-int>' with token 'none', tabfield, or tabfield-suffix."
      }
    }

    if ($keyMap.ContainsKey("output_map")) {
      $outputMapValue = [string]$keyMap["output_map"][0].value
      if (-not (Test-OutputMapValue -Value $outputMapValue)) {
        Add-Issue -Issues $issues -Code "INVALID_OUTPUT_MAP" -Line $keyMap["output_map"][0].line -Section $section.name -Key "output_map" -Message "output_map must be empty or comma-separated '<tabfield>:<column>:<row>' entries."
      }
    }

    if ($keyMap.ContainsKey("config_id")) {
      $templateName = $section.name.Substring("TEMPLATE:".Length)
      $configIdValue = [string]$keyMap["config_id"][0].value
      if (-not [string]::IsNullOrWhiteSpace($templateName) -and -not [string]::IsNullOrWhiteSpace($configIdValue)) {
        $pair = "$templateName|$configIdValue"
        if (-not $templateConfigIds.Add($pair)) {
          $duplicateTemplateConfigIdCount++
          Add-Issue -Issues $issues -Code "DUPLICATE_TEMPLATE_CONFIG_ID" -Line $section.header_line -Section $section.name -Key "config_id" -Message "Duplicate template/config_id pair '$pair' is not allowed."
        }
      }
    }
  }
  elseif (-not $allowedNonTemplateSections.Contains($section.name)) {
    Add-Issue -Issues $issues -Code "UNEXPECTED_SECTION" -Line $section.header_line -Section $section.name -Key "" -Message "Section '$($section.name)' is not part of contract 1.0.0 for TemplateConfig."
  }
}

$validationPass = ($issues.Count -eq 0)
$result = [ordered]@{
  schema_version = 1
  validator = $validatorId
  contract_version = $contractVersion
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
  template_section_count = $templateSectionCount
  duplicate_template_config_id_count = $duplicateTemplateConfigIdCount
  issue_count = $issues.Count
  valid = $validationPass
  required_keys = $requiredKeys
  allowed_empty_keys = $allowedEmptyKeys
  issues = $issues.ToArray()
}

$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding ASCII

$txtLines = [System.Collections.Generic.List[string]]::new()
$txtLines.Add("VALIDATOR=$validatorId") | Out-Null
$txtLines.Add("CONTRACT_VERSION=$contractVersion") | Out-Null
$txtLines.Add("RUN_LABEL=$runLabelFinal") | Out-Null
$txtLines.Add("CASE_LABEL=$CaseLabel") | Out-Null
$txtLines.Add("TARGET_PATH=$($parseResult.path)") | Out-Null
$txtLines.Add("PARSEABLE=$($parseResult.parseable)") | Out-Null
$txtLines.Add("SECTION_COUNT=$($parseResult.section_count)") | Out-Null
$txtLines.Add("TEMPLATE_SECTION_COUNT=$templateSectionCount") | Out-Null
$txtLines.Add("ISSUE_COUNT=$($issues.Count)") | Out-Null
$txtLines.Add("RESULT=$(if ($validationPass) { 'PASS' } else { 'FAIL' })") | Out-Null
foreach ($issue in $issues) {
  $txtLines.Add("ISSUE LINE=$($issue.line) CODE=$($issue.code) SECTION=$($issue.section) KEY=$($issue.key) MSG=$($issue.message)") | Out-Null
}
$txtLines | Set-Content -LiteralPath $txtPath -Encoding ASCII

Write-Host "VALIDATOR=$validatorId"
Write-Host "CONTRACT_VERSION=$contractVersion"
Write-Host "RUN_LABEL=$runLabelFinal"
Write-Host "CASE_LABEL=$CaseLabel"
Write-Host "TARGET_PATH=$($parseResult.path)"
Write-Host "PARSEABLE=$($parseResult.parseable)"
Write-Host "SECTION_COUNT=$($parseResult.section_count)"
Write-Host "TEMPLATE_SECTION_COUNT=$templateSectionCount"
Write-Host "ISSUE_COUNT=$($issues.Count)"
Write-Host ("RESULT={0}" -f ($(if ($validationPass) { "PASS" } else { "FAIL" })))
Write-Host "JSON_SUMMARY=$jsonPath"
Write-Host "TXT_SUMMARY=$txtPath"

if ($validationPass) {
  exit 0
}

exit 2
