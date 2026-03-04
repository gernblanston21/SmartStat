param(
  [Parameter(Mandatory = $true)][string]$IniPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-KeySet {
  return ,([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase))
}

function Add-Issue {
  param(
    [Parameter(Mandatory = $true)][object]$Issues,
    [Parameter(Mandatory = $true)][string]$Code,
    [Parameter(Mandatory = $true)][int]$Line,
    [Parameter(Mandatory = $true)][string]$Message,
    [Parameter(Mandatory = $false)][bool]$IsPartialWrite = $false
  )

  $Issues.Add([pscustomobject]@{
      code = $Code
      line = $Line
      message = $Message
      partial_write_indicator = $IsPartialWrite
    }) | Out-Null
}

function Validate-Ini {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $issues = [System.Collections.Generic.List[object]]::new()
  $resolvedPath = ""
  $lines = @()

  if (-not (Test-Path -LiteralPath $Path)) {
    Add-Issue -Issues $issues -Code "FILE_NOT_FOUND" -Line 0 -Message "INI path not found: $Path" -IsPartialWrite:$false
    return [pscustomobject]@{
      path = $Path
      exists = $false
      parseable = $false
      valid = $false
      section_count = 0
      key_count = 0
      duplicate_key_count = 0
      partial_write_indicator_count = 0
      issue_count = $issues.Count
      issues = $issues.ToArray()
    }
  }

  $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
  $lines = Get-Content -LiteralPath $resolvedPath

  $currentSection = ""
  $sectionKeys = @{}
  $sectionSeen = New-KeySet
  $sectionCount = 0
  $keyCount = 0
  $duplicateCount = 0
  $parseErrorCount = 0

  for ($i = 0; $i -lt $lines.Count; $i++) {
    $lineNumber = $i + 1
    $line = [string]$lines[$i]
    $trimmed = $line.Trim()

    if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
    if ($trimmed.StartsWith(";") -or $trimmed.StartsWith("#")) { continue }

    if ($trimmed -match '^(<<<<<<<|=======|>>>>>>>)') {
      Add-Issue -Issues $issues -Code "PARTIAL_WRITE_MARKER" -Line $lineNumber -Message "Conflict/partial marker detected." -IsPartialWrite:$true
      $parseErrorCount++
      continue
    }

    if ($trimmed.StartsWith("[")) {
      if (-not $trimmed.EndsWith("]")) {
        Add-Issue -Issues $issues -Code "MALFORMED_SECTION_HEADER" -Line $lineNumber -Message "Section header must end with ']': $trimmed" -IsPartialWrite:$true
        $parseErrorCount++
        continue
      }

      $inner = $trimmed.Substring(1, $trimmed.Length - 2).Trim()
      if ([string]::IsNullOrWhiteSpace($inner)) {
        Add-Issue -Issues $issues -Code "EMPTY_SECTION_NAME" -Line $lineNumber -Message "Section name cannot be empty." -IsPartialWrite:$true
        $parseErrorCount++
        continue
      }

      $currentSection = $inner
      if (-not $sectionSeen.Contains($currentSection)) {
        $sectionSeen.Add($currentSection) | Out-Null
        $sectionCount++
      }

      if (-not $sectionKeys.ContainsKey($currentSection)) {
        $sectionKeys[$currentSection] = New-KeySet
      }
      continue
    }

    if ([string]::IsNullOrWhiteSpace($currentSection)) {
      Add-Issue -Issues $issues -Code "KEY_OUTSIDE_SECTION" -Line $lineNumber -Message "Key/value appears before any section header." -IsPartialWrite:$true
      $parseErrorCount++
      continue
    }

    $eqIndex = $line.IndexOf("=")
    if ($eqIndex -lt 0) {
      Add-Issue -Issues $issues -Code "MISSING_EQUALS" -Line $lineNumber -Message "Expected key=value format." -IsPartialWrite:$true
      $parseErrorCount++
      continue
    }

    $key = $line.Substring(0, $eqIndex).Trim()
    $value = $line.Substring($eqIndex + 1).Trim()

    if ([string]::IsNullOrWhiteSpace($key)) {
      Add-Issue -Issues $issues -Code "EMPTY_KEY" -Line $lineNumber -Message "Key name cannot be empty." -IsPartialWrite:$true
      $parseErrorCount++
      continue
    }

    $keysForSection = $sectionKeys[$currentSection]
    if ($keysForSection.Contains($key)) {
      Add-Issue -Issues $issues -Code "DUPLICATE_KEY" -Line $lineNumber -Message "Duplicate key '$key' in section '$currentSection'." -IsPartialWrite:$false
      $duplicateCount++
      continue
    }

    $keysForSection.Add($key) | Out-Null
    $keyCount++
  }

  $partialCount = 0
  foreach ($issue in $issues) {
    if ([bool]$issue.partial_write_indicator) {
      $partialCount++
    }
  }

  $parseable = ($parseErrorCount -eq 0)
  $valid = ($issues.Count -eq 0)

  return [pscustomobject]@{
    path = $resolvedPath
    exists = $true
    parseable = $parseable
    valid = $valid
    section_count = $sectionCount
    key_count = $keyCount
    duplicate_key_count = $duplicateCount
    partial_write_indicator_count = $partialCount
    issue_count = $issues.Count
    issues = $issues.ToArray()
  }
}

$result = Validate-Ini -Path $IniPath

Write-Host "VALIDATION_PATH=$($result.path)"
Write-Host "VALIDATION_EXISTS=$($result.exists)"
Write-Host "VALIDATION_PARSEABLE=$($result.parseable)"
Write-Host "VALIDATION_VALID=$($result.valid)"
Write-Host "SECTION_COUNT=$($result.section_count)"
Write-Host "KEY_COUNT=$($result.key_count)"
Write-Host "DUPLICATE_KEY_COUNT=$($result.duplicate_key_count)"
Write-Host "PARTIAL_WRITE_INDICATOR_COUNT=$($result.partial_write_indicator_count)"
Write-Host "ISSUE_COUNT=$($result.issue_count)"

if ($result.issue_count -gt 0) {
  Write-Host "ISSUES_BEGIN"
  foreach ($issue in $result.issues) {
    Write-Host ("LINE={0} CODE={1} PARTIAL={2} MSG={3}" -f $issue.line, $issue.code, $issue.partial_write_indicator, $issue.message)
  }
  Write-Host "ISSUES_END"
}

Write-Host "JSON_BEGIN"
$result | ConvertTo-Json -Depth 8
Write-Host "JSON_END"

if ($result.valid) {
  exit 0
}

exit 2
