param(
  [Parameter(Mandatory = $false)][string]$ManifestPath = "",
  [Parameter(Mandatory = $false)][string]$OutputRoot = "",
  [Parameter(Mandatory = $false)][string]$RunLabel = "",
  [Parameter(Mandatory = $false)][switch]$StopOnFailure
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

function Get-OptionalString {
  param(
    [Parameter(Mandatory = $true)][object]$Object,
    [Parameter(Mandatory = $true)][string]$Name
  )

  if (-not (Test-HasProperty -Object $Object -Name $Name)) {
    return ""
  }

  return ([string]$Object.$Name).Trim()
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

function Resolve-RepoPath {
  param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$ManifestDir,
    [Parameter(Mandatory = $true)][string]$RepoRoot
  )

  $p = $InputPath.Trim()
  if (-not [System.IO.Path]::IsPathRooted($p)) {
    $p = $p -replace '/', '\'
  }

  if ([System.IO.Path]::IsPathRooted($p)) {
    if (-not (Test-Path -LiteralPath $p)) {
      throw "Path not found: $p"
    }
    return (Resolve-Path -LiteralPath $p).Path
  }

  $candidateManifest = Join-Path $ManifestDir $p
  if (Test-Path -LiteralPath $candidateManifest) {
    return (Resolve-Path -LiteralPath $candidateManifest).Path
  }

  $candidateRepo = Join-Path $RepoRoot $p
  if (Test-Path -LiteralPath $candidateRepo) {
    return (Resolve-Path -LiteralPath $candidateRepo).Path
  }

  throw "Relative path not found from manifest or repo root: $p"
}

function Normalize-Args {
  param(
    [Parameter(Mandatory = $false)][object]$ArgsValue
  )

  $argsOut = @()
  if ($null -eq $ArgsValue) {
    return $argsOut
  }

  foreach ($arg in @($ArgsValue)) {
    $argsOut += ([string]$arg)
  }

  return $argsOut
}

$cscript = Get-Command cscript.exe -ErrorAction SilentlyContinue
if (-not $cscript) {
  throw "cscript.exe was not found on PATH."
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
  throw "Manifest file not found: $ManifestPath"
}

$manifestFullPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDir = Split-Path -Parent $manifestFullPath
$repoRoot = Find-RepoRoot -startDir $PSScriptRoot

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

$defaultScript = ""
if ((Test-HasProperty -Object $manifest -Name "defaults") -and $null -ne $manifest.defaults) {
  $defaultScript = Get-OptionalString -Object $manifest.defaults -Name "script"
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

if ([string]::IsNullOrWhiteSpace($RunLabel)) {
  $RunLabel = "run_" + [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")
}

$outputRootFull = ""
if ([System.IO.Path]::IsPathRooted($OutputRoot)) {
  $outputRootFull = [System.IO.Path]::GetFullPath($OutputRoot)
}
else {
  $outputRootFull = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputRoot))
}

$runDir = Join-Path $outputRootFull $RunLabel
$caseDir = Join-Path $runDir "cases"
New-Item -ItemType Directory -Force -Path $caseDir | Out-Null

$startedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")

$results = New-Object System.Collections.Generic.List[object]
$failureCount = 0
$seenCaseIds = @{}

foreach ($case in $enabledCases) {
  $caseId = Get-RequiredString -Object $case -Name "id" -Context "case"
  if ($seenCaseIds.ContainsKey($caseId)) {
    throw "Duplicate case id found: $caseId"
  }
  $seenCaseIds[$caseId] = $true

  $runnerRel = Get-RequiredString -Object $case -Name "runner" -Context "case '$caseId'"
  $runnerPath = Resolve-RepoPath -InputPath $runnerRel -ManifestDir $manifestDir -RepoRoot $repoRoot

  $scriptRel = Get-OptionalString -Object $case -Name "script"
  if ([string]::IsNullOrWhiteSpace($scriptRel)) {
    if ([string]::IsNullOrWhiteSpace($defaultScript)) {
      throw "Case '$caseId' has no script and manifest.defaults.script is empty."
    }
    $scriptRel = $defaultScript
  }
  $scriptPath = Resolve-RepoPath -InputPath $scriptRel -ManifestDir $manifestDir -RepoRoot $repoRoot

  $artifactRel = Get-RequiredString -Object $case -Name "artifact" -Context "case '$caseId'"
  if ([System.IO.Path]::IsPathRooted($artifactRel)) {
    throw "case '$caseId' has rooted artifact path; only relative paths are allowed."
  }
  $artifactPath = Join-Path $caseDir $artifactRel
  $artifactParent = Split-Path -Parent $artifactPath
  if (-not [string]::IsNullOrWhiteSpace($artifactParent)) {
    New-Item -ItemType Directory -Force -Path $artifactParent | Out-Null
  }

  $caseArgs = @(
    Normalize-Args -ArgsValue $case.args
  )
  $invokeArgs = New-Object System.Collections.Generic.List[string]
  $invokeArgs.Add("//nologo") | Out-Null
  $invokeArgs.Add($runnerPath) | Out-Null
  $invokeArgs.Add($scriptPath) | Out-Null
  foreach ($arg in $caseArgs) {
    $invokeArgs.Add($arg) | Out-Null
  }

  $rawOutput = & cscript.exe $invokeArgs.ToArray() 2>&1
  $exitCode = $LASTEXITCODE

  if ($null -eq $rawOutput) {
    "" | Set-Content -LiteralPath $artifactPath -Encoding ASCII
  }
  else {
    ($rawOutput | ForEach-Object { [string]$_ }) | Set-Content -LiteralPath $artifactPath -Encoding ASCII
  }

  $artifactHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $artifactPath).Hash

  if ($exitCode -ne 0) {
    $failureCount++
  }

  $results.Add([pscustomobject]@{
      case_id = $caseId
      runner = $runnerRel
      script = $scriptRel
      args = $caseArgs
      artifact = $artifactRel
      artifact_sha256 = $artifactHash
      exit_code = $exitCode
    }) | Out-Null

  if ($StopOnFailure -and $exitCode -ne 0) {
    break
  }
}

$endedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
$summaryJsonPath = Join-Path $runDir "run_summary.json"
$summaryTxtPath = Join-Path $runDir "run_summary.txt"

$summary = [ordered]@{
  schema_version = 1
  pack_name = $packName
  manifest_path = $manifestFullPath
  run_label = $RunLabel
  started_utc = $startedUtc
  ended_utc = $endedUtc
  case_count = $results.Count
  failure_count = $failureCount
  success = ($failureCount -eq 0)
  results = $results.ToArray()
}

$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJsonPath -Encoding ASCII

$txt = New-Object System.Collections.Generic.List[string]
$txt.Add("PACK_NAME=$packName") | Out-Null
$txt.Add("RUN_LABEL=$RunLabel") | Out-Null
$txt.Add("STARTED_UTC=$startedUtc") | Out-Null
$txt.Add("ENDED_UTC=$endedUtc") | Out-Null
$txt.Add("CASE_COUNT=$($results.Count)") | Out-Null
$txt.Add("FAILURE_COUNT=$failureCount") | Out-Null
$txt.Add("RUN_PASS=$($failureCount -eq 0)") | Out-Null
foreach ($r in $results) {
  $txt.Add("CASE=$($r.case_id) EXIT_CODE=$($r.exit_code) SHA256=$($r.artifact_sha256) ARTIFACT=$($r.artifact)") | Out-Null
}
$txt | Set-Content -LiteralPath $summaryTxtPath -Encoding ASCII

Write-Host "RUN_LABEL=$RunLabel"
Write-Host "RUN_DIR=$runDir"
Write-Host "SUMMARY_JSON=$summaryJsonPath"
Write-Host "SUMMARY_TXT=$summaryTxtPath"

if ($failureCount -ne 0) {
  exit 2
}

exit 0
