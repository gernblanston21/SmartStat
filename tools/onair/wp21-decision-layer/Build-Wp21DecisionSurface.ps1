param(
  [Parameter(Mandatory = $true)]
  [string]$InputArtifactPath,

  [Parameter(Mandatory = $true)]
  [string]$OutputArtifactPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$reasonToMessage = [ordered]@{
  INELIGIBLE_EVIDENCE_PRESENT = "Preview is not runtime-eligible. Do not proceed."
  ERRORS_PRESENT = "Blocking errors are present in preview evidence."
  REQUIRED_INPUT_MISSING_OR_MALFORMED = "Required preview evidence is missing or malformed."
  AMBIGUOUS_INPUT_SHAPE = "Preview input shape is ambiguous and cannot be classified safely."
  WARNINGS_PRESENT = "Warnings are present. Review preview evidence before proceeding."
  NON_PASS_RULE_OUTCOME_PRESENT = "One or more rule outcomes are not PASS. Review evidence."
  NO_BLOCKERS_OR_REVIEW_SIGNALS = "No blockers or review signals detected in preview evidence."
}

function Test-IsMap {
  param([object]$Value)
  return ($Value -is [System.Collections.IDictionary])
}

function Test-IsArrayLike {
  param([object]$Value)
  if ($null -eq $Value) {
    return $false
  }
  if ($Value -is [string]) {
    return $false
  }
  return (
    ($Value -is [System.Array]) -or
    ($Value -is [System.Collections.IList])
  )
}

function Convert-JsonNodeToDeterministicValue {
  param([object]$Value)

  if ($null -eq $Value) {
    return $null
  }

  if ($Value -is [System.Collections.IDictionary]) {
    $mapped = [ordered]@{}
    foreach ($key in $Value.Keys) {
      $mapped[[string]$key] = Convert-JsonNodeToDeterministicValue -Value $Value[$key]
    }
    return $mapped
  }

  if ($Value -is [System.Management.Automation.PSCustomObject]) {
    $mapped = [ordered]@{}
    foreach ($property in $Value.PSObject.Properties) {
      $mapped[$property.Name] = Convert-JsonNodeToDeterministicValue -Value $property.Value
    }
    return $mapped
  }

  if (Test-IsArrayLike -Value $Value) {
    $items = @()
    foreach ($item in @($Value)) {
      $items += ,(Convert-JsonNodeToDeterministicValue -Value $item)
    }
    return ,([object[]]$items)
  }

  return $Value
}

function Get-MapMemberOrNull {
  param(
    [System.Collections.IDictionary]$Map,
    [string]$Key
  )

  if ($null -eq $Map) {
    return $null
  }
  if (-not $Map.Contains($Key)) {
    return $null
  }
  $candidate = $Map[$Key]
  if (-not (Test-IsMap -Value $candidate)) {
    return $null
  }
  return $candidate
}

function Get-ArrayMemberOrNull {
  param(
    [System.Collections.IDictionary]$Map,
    [string]$Key
  )

  if ($null -eq $Map) {
    return $null
  }
  if (-not $Map.Contains($Key)) {
    return $null
  }
  $candidate = $Map[$Key]
  if (-not (Test-IsArrayLike -Value $candidate)) {
    return $null
  }
  return ,([object[]]$candidate)
}

function New-DecisionFromReason {
  param([string]$Reason)

  if (-not $reasonToMessage.Contains($Reason)) {
    throw "unsupported decision reason: $Reason"
  }

  $status = ""
  $action = ""

  switch ($Reason) {
    "INELIGIBLE_EVIDENCE_PRESENT" {
      $status = "BLOCKED"
      $action = "DO_NOT_PROCEED_ESCALATE"
    }
    "ERRORS_PRESENT" {
      $status = "BLOCKED"
      $action = "DO_NOT_PROCEED_ESCALATE"
    }
    "REQUIRED_INPUT_MISSING_OR_MALFORMED" {
      $status = "BLOCKED"
      $action = "DO_NOT_PROCEED_ESCALATE"
    }
    "AMBIGUOUS_INPUT_SHAPE" {
      $status = "BLOCKED"
      $action = "DO_NOT_PROCEED_ESCALATE"
    }
    "WARNINGS_PRESENT" {
      $status = "REVIEW_REQUIRED"
      $action = "REVIEW_PREVIEW_EVIDENCE"
    }
    "NON_PASS_RULE_OUTCOME_PRESENT" {
      $status = "REVIEW_REQUIRED"
      $action = "REVIEW_PREVIEW_EVIDENCE"
    }
    "NO_BLOCKERS_OR_REVIEW_SIGNALS" {
      $status = "AUTO_SAFE"
      $action = "PROCEED_WITH_OPERATOR_FLOW"
    }
    default {
      throw "unsupported decision reason: $Reason"
    }
  }

  return [ordered]@{
    decision_status = $status
    decision_reason = $Reason
    operator_message = $reasonToMessage[$Reason]
    recommended_action = $action
  }
}

function Get-EligibleBranchContextOrNull {
  param([System.Collections.IDictionary]$PreviewPayload)

  $projectionMetadata = Get-MapMemberOrNull -Map $PreviewPayload -Key "projection_metadata"
  $ruleSummaryPreview = Get-MapMemberOrNull -Map $PreviewPayload -Key "rule_evaluation_summary_preview"
  $tracePreview = Get-MapMemberOrNull -Map $PreviewPayload -Key "rule_evaluation_trace_preview"

  if ($null -eq $projectionMetadata -or $null -eq $ruleSummaryPreview -or $null -eq $tracePreview) {
    return $null
  }

  return [ordered]@{
    projection_metadata = $projectionMetadata
    rule_evaluation_summary_preview = $ruleSummaryPreview
    rule_evaluation_trace_preview = $tracePreview
  }
}

function Get-NonEligibleBranchContextOrNull {
  param([System.Collections.IDictionary]$PreviewPayload)

  $ineligibleEvidencePreview = Get-MapMemberOrNull -Map $PreviewPayload -Key "ineligible_evidence_preview"
  if ($null -eq $ineligibleEvidencePreview) {
    return $null
  }

  return [ordered]@{
    ineligible_evidence_preview = $ineligibleEvidencePreview
  }
}

function Test-OrderedRulesMalformed {
  param([object[]]$OrderedRules)

  foreach ($entry in $OrderedRules) {
    if (-not (Test-IsMap -Value $entry)) {
      return $true
    }
    if (-not $entry.Contains("outcome")) {
      return $true
    }
    if ($entry["outcome"] -isnot [string]) {
      return $true
    }
    if ([string]::IsNullOrWhiteSpace([string]$entry["outcome"])) {
      return $true
    }
  }

  return $false
}

function Test-AnyNonPassOutcome {
  param([object[]]$OrderedRules)

  foreach ($entry in $OrderedRules) {
    $outcome = [string]$entry["outcome"]
    if ($outcome -cne "PASS") {
      return $true
    }
  }

  return $false
}

function Get-DecisionReason {
  param([System.Collections.IDictionary]$InputArtifact)

  $previewPayload = Get-MapMemberOrNull -Map $InputArtifact -Key "preview_payload"
  if ($null -eq $previewPayload) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  $eligibleContext = Get-EligibleBranchContextOrNull -PreviewPayload $previewPayload
  $nonEligibleContext = Get-NonEligibleBranchContextOrNull -PreviewPayload $previewPayload

  $hasEligibleShape = ($null -ne $eligibleContext)
  $hasNonEligibleShape = ($null -ne $nonEligibleContext)

  if ($hasEligibleShape -and $hasNonEligibleShape) {
    return "AMBIGUOUS_INPUT_SHAPE"
  }

  if ($hasNonEligibleShape) {
    return "INELIGIBLE_EVIDENCE_PRESENT"
  }

  if (-not $hasEligibleShape) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  $issuesSummary = Get-MapMemberOrNull -Map $eligibleContext["projection_metadata"] -Key "issues_summary"
  if ($null -eq $issuesSummary) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  $errors = Get-ArrayMemberOrNull -Map $issuesSummary -Key "errors"
  if ($null -eq $errors) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  if ($errors.Count -gt 0) {
    return "ERRORS_PRESENT"
  }

  $warnings = Get-ArrayMemberOrNull -Map $issuesSummary -Key "warnings"
  $orderedRules = Get-ArrayMemberOrNull -Map $eligibleContext["rule_evaluation_summary_preview"] -Key "ordered_rules"
  if ($null -eq $warnings -or $null -eq $orderedRules) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  if (Test-OrderedRulesMalformed -OrderedRules $orderedRules) {
    return "REQUIRED_INPUT_MISSING_OR_MALFORMED"
  }

  if ($warnings.Count -gt 0) {
    return "WARNINGS_PRESENT"
  }

  if (Test-AnyNonPassOutcome -OrderedRules $orderedRules) {
    return "NON_PASS_RULE_OUTCOME_PRESENT"
  }

  return "NO_BLOCKERS_OR_REVIEW_SIGNALS"
}

function Parse-JsonToDeterministicRootObject {
  param([string]$RawJson)

  Add-Type -AssemblyName System.Web.Extensions
  $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $serializer.MaxJsonLength = [int]::MaxValue

  $parsed = $serializer.DeserializeObject($RawJson)
  $normalized = Convert-JsonNodeToDeterministicValue -Value $parsed

  if (-not (Test-IsMap -Value $normalized)) {
    throw "input root must be a JSON object"
  }

  return $normalized
}

function Load-InputArtifact {
  param([string]$Path)

  $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
  $rawJson = Get-Content -Raw -LiteralPath $resolvedPath
  return Parse-JsonToDeterministicRootObject -RawJson $rawJson
}

try {
  $inputArtifact = Load-InputArtifact -Path $InputArtifactPath
} catch {
  Write-Error ("TRANSPORT_READ_FAILURE: " + $_.Exception.Message)
  exit 1
}

$reason = Get-DecisionReason -InputArtifact $inputArtifact
$decisionArtifact = New-DecisionFromReason -Reason $reason

try {
  $outputDir = Split-Path -Parent $OutputArtifactPath
  if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
  }

  $decisionJson = $decisionArtifact | ConvertTo-Json -Depth 5
  Set-Content -LiteralPath $OutputArtifactPath -Value $decisionJson -Encoding UTF8
} catch {
  Write-Error ("OUTPUT_WRITE_FAILURE: " + $_.Exception.Message)
  exit 1
}
