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

function Resolve-InputPath {
  param(
    [Parameter(Mandatory = $true)][string]$RepoRoot,
    [Parameter(Mandatory = $false)][string]$Path,
    [Parameter(Mandatory = $true)][string]$DefaultFileName
  )

  $candidate = ""
  if ([string]::IsNullOrWhiteSpace($Path)) {
    $candidate = Join-Path $RepoRoot $DefaultFileName
  }
  elseif ([System.IO.Path]::IsPathRooted($Path)) {
    $candidate = $Path
  }
  else {
    $candidate = Join-Path $RepoRoot $Path
  }

  return [System.IO.Path]::GetFullPath($candidate)
}

function New-RunLabel {
  param(
    [Parameter(Mandatory = $false)][string]$RunLabel
  )

  if ([string]::IsNullOrWhiteSpace($RunLabel)) {
    return "run_" + [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")
  }

  return ($RunLabel.Trim() -replace '[^A-Za-z0-9._-]', '_')
}

function Get-SafeToken {
  param(
    [Parameter(Mandatory = $false)][string]$Value,
    [Parameter(Mandatory = $false)][string]$Fallback = "case"
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return $Fallback
  }

  $clean = ($Value.Trim() -replace '[^A-Za-z0-9._-]', '_')
  if ([string]::IsNullOrWhiteSpace($clean)) {
    return $Fallback
  }

  return $clean
}

function New-IssueList {
  return ,([System.Collections.Generic.List[object]]::new())
}

function Add-Issue {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)][string]$Code,
    [Parameter(Mandatory = $false)][int]$Line = 0,
    [Parameter(Mandatory = $false)][string]$Section = "",
    [Parameter(Mandatory = $false)][string]$Key = "",
    [Parameter(Mandatory = $true)][string]$Message
  )

  $Issues.Add([pscustomobject]@{
      code = $Code
      line = $Line
      section = $Section
      key = $Key
      message = $Message
    }) | Out-Null
}

function Get-LineEndingStyle {
  param(
    [Parameter(Mandatory = $true)][string]$Text
  )

  $hasCRLF = $Text.Contains("`r`n")
  $hasLFOnly = [System.Text.RegularExpressions.Regex]::IsMatch($Text, "(?<!\r)\n")

  if ($hasCRLF -and $hasLFOnly) { return "mixed" }
  if ($hasCRLF) { return "crlf" }
  if ($hasLFOnly) { return "lf" }
  return "none"
}

function Parse-IniDocument {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $issues = New-IssueList

  try {
    if (-not (Test-Path -LiteralPath $Path)) {
      Add-Issue -Issues $issues -Code "FILE_NOT_FOUND" -Line 0 -Section "" -Key "" -Message "INI path not found: $Path"
      return [pscustomobject]@{
        path = $Path
        exists = $false
        utf8_bom = $false
        line_ending = "none"
        line_count = 0
        section_count = 0
        key_count = 0
        parseable = $false
        parse_issue_count = 1
        sections = @()
        issues = $issues.ToArray()
      }
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
    $utf8Bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $rawText = [string]([System.IO.File]::ReadAllText($resolvedPath))
    $lineEnding = Get-LineEndingStyle -Text $rawText
    $lines = @(Get-Content -LiteralPath $resolvedPath)

    $sectionsInternal = [System.Collections.Generic.List[object]]::new()
    $currentSection = $null
    $parseIssueCount = 0

    for ($i = 0; $i -lt $lines.Count; $i++) {
      $lineNumber = $i + 1
      $line = [string]$lines[$i]
      $trimmed = $line.Trim()

      if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
      if ($trimmed.StartsWith(";") -or $trimmed.StartsWith("#") -or $trimmed.StartsWith("'")) { continue }

      if ($trimmed.StartsWith("[")) {
        if (-not $trimmed.EndsWith("]")) {
          Add-Issue -Issues $issues -Code "MALFORMED_SECTION_HEADER" -Line $lineNumber -Section "" -Key "" -Message "Section header must end with ']': $trimmed"
          $parseIssueCount++
          $currentSection = $null
          continue
        }

        $sectionName = $trimmed.Substring(1, $trimmed.Length - 2).Trim()
        if ([string]::IsNullOrWhiteSpace($sectionName)) {
          Add-Issue -Issues $issues -Code "EMPTY_SECTION_NAME" -Line $lineNumber -Section "" -Key "" -Message "Section name cannot be empty."
          $parseIssueCount++
          $currentSection = $null
          continue
        }

        $currentSection = [pscustomobject]@{
          name = $sectionName
          header_line = $lineNumber
          header_text = $trimmed
          keys = [System.Collections.Generic.List[object]]::new()
        }
        $sectionsInternal.Add($currentSection) | Out-Null
        continue
      }

      if ($null -eq $currentSection) {
        Add-Issue -Issues $issues -Code "KEY_OUTSIDE_SECTION" -Line $lineNumber -Section "" -Key "" -Message "Key/value appears before any valid section header."
        $parseIssueCount++
        continue
      }

      $eqIndex = $line.IndexOf("=")
      if ($eqIndex -lt 0) {
        Add-Issue -Issues $issues -Code "MISSING_EQUALS" -Line $lineNumber -Section $currentSection.name -Key "" -Message "Expected key=value format."
        $parseIssueCount++
        continue
      }

      $key = $line.Substring(0, $eqIndex).Trim()
      $value = $line.Substring($eqIndex + 1).Trim()

      if ([string]::IsNullOrWhiteSpace($key)) {
        Add-Issue -Issues $issues -Code "EMPTY_KEY" -Line $lineNumber -Section $currentSection.name -Key "" -Message "Key name cannot be empty."
        $parseIssueCount++
        continue
      }

      $currentSection.keys.Add([pscustomobject]@{
          key = $key
          key_normalized = $key.ToLowerInvariant()
          value = $value
          line = $lineNumber
          raw = $line
        }) | Out-Null
    }

    $sectionsOutput = [System.Collections.Generic.List[object]]::new()
    $keyCount = 0
    foreach ($section in $sectionsInternal) {
      $keysArray = $section.keys.ToArray()
      $keyCount += $keysArray.Count
      $sectionsOutput.Add([pscustomobject]@{
          name = [string]$section.name
          header_line = [int]$section.header_line
          header_text = [string]$section.header_text
          keys = $keysArray
        }) | Out-Null
    }

    return [pscustomobject]@{
      path = $resolvedPath
      exists = $true
      utf8_bom = $utf8Bom
      line_ending = $lineEnding
      line_count = $lines.Count
      section_count = $sectionsOutput.Count
      key_count = $keyCount
      parseable = ($parseIssueCount -eq 0)
      parse_issue_count = $parseIssueCount
      sections = $sectionsOutput.ToArray()
      issues = $issues.ToArray()
    }
  }
  catch {
    throw "Parse-IniDocument internal failure at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
  }
}
