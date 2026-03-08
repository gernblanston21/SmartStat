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
  $list = [System.Collections.Generic.List[object]]::new()
  Write-Output -NoEnumerate $list
}

function Add-Issue {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)][string]$Code,
    [Parameter(Mandatory = $true)][string]$Message,
    [Parameter(Mandatory = $false)][string]$Path = ""
  )

  $Issues.Add([pscustomobject]@{
      code = $Code
      path = $Path
      message = $Message
    }) | Out-Null
}

function Get-CanonicalJson {
  param(
    [Parameter(Mandatory = $true)]$Value
  )

  return ($Value | ConvertTo-Json -Depth 100 -Compress)
}

function Get-Sha256Hex {
  param(
    [Parameter(Mandatory = $true)][string]$Text
  )

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hashBytes = $sha.ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hashBytes)).Replace("-", "")
  }
  finally {
    $sha.Dispose()
  }
}
