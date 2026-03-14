param(
  [string]$PositiveRun1 = "tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json",
  [string]$PositiveRun2 = "tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run2.json",
  [string]$ReportOut = "tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/traceability_subtree_equivalence_report.json"
)

$ErrorActionPreference = "Stop"

$overlapKeys = @(
  "projection_contract",
  "projection_kind",
  "input_artifact",
  "input_identity",
  "deterministic_identity_summary"
)

$expectedProjectionMetadataKeyOrder = @(
  "projection_contract",
  "projection_kind",
  "input_artifact",
  "input_identity",
  "status_summary",
  "deterministic_identity_summary",
  "semantic_interpretation_summary",
  "issues_summary"
)

function Get-Sha256Hex {
  param([Parameter(Mandatory = $true)][string]$Text)

  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hashBytes = $sha.ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hashBytes)).Replace("-", "")
  } finally {
    $sha.Dispose()
  }
}

function Load-OrderedJson {
  param([Parameter(Mandatory = $true)][string]$Path)

  $resolved = (Resolve-Path $Path).Path
  $raw = Get-Content -Raw -Path $resolved
  return ($raw | ConvertFrom-Json -AsHashtable)
}

function Build-BoundedOverlap {
  param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Source)

  $overlap = [ordered]@{}
  foreach ($key in $overlapKeys) {
    if (-not $Source.Contains($key)) {
      throw "required overlap key missing: $key"
    }
    $overlap[$key] = $Source[$key]
  }
  return $overlap
}

function Build-RunEvidence {
  param([Parameter(Mandatory = $true)][string]$Path)

  $resolved = (Resolve-Path $Path).Path
  $doc = Load-OrderedJson -Path $resolved
  $previewPayload = $doc["preview_payload"]

  if ($null -eq $previewPayload) {
    throw "preview_payload missing in $resolved"
  }

  $projectionMetadata = $previewPayload["projection_metadata"]
  $tracePreview = $previewPayload["rule_evaluation_trace_preview"]

  if ($null -eq $projectionMetadata) {
    throw "projection_metadata missing in $resolved"
  }
  if ($null -eq $tracePreview) {
    throw "rule_evaluation_trace_preview missing in $resolved"
  }

  $projectionMetadataKeysActual = @($projectionMetadata.Keys)
  $projectionMetadataOrderExpected = @($expectedProjectionMetadataKeyOrder)
  $projectionMetadataOrderMatch = (($projectionMetadataKeysActual -join "|") -ceq ($projectionMetadataOrderExpected -join "|"))

  $projectionOverlap = Build-BoundedOverlap -Source $projectionMetadata
  $traceOverlap = Build-BoundedOverlap -Source $tracePreview

  $projectionOverlapJson = $projectionOverlap | ConvertTo-Json -Depth 20 -Compress
  $traceOverlapJson = $traceOverlap | ConvertTo-Json -Depth 20 -Compress

  return [ordered]@{
    artifact_path = $resolved
    overlap_field_set = @($overlapKeys)
    projection_metadata_overlap_sha256 = (Get-Sha256Hex -Text $projectionOverlapJson)
    rule_evaluation_trace_preview_overlap_sha256 = (Get-Sha256Hex -Text $traceOverlapJson)
    overlap_byte_identical = ($projectionOverlapJson -ceq $traceOverlapJson)
    projection_metadata_key_order_expected = @($projectionMetadataOrderExpected)
    projection_metadata_key_order_actual = @($projectionMetadataKeysActual)
    projection_metadata_key_order_match = $projectionMetadataOrderMatch
  }
}

$run1 = Build-RunEvidence -Path $PositiveRun1
$run2 = Build-RunEvidence -Path $PositiveRun2

$report = [ordered]@{
  report_kind = "slice02_traceability_subtree_equivalence_v1"
  comparison_mode = "bounded_overlap_only"
  positive_runs = [ordered]@{
    run1 = $run1
    run2 = $run2
  }
  cross_run = [ordered]@{
    projection_metadata_overlap_sha256_match = ($run1["projection_metadata_overlap_sha256"] -ceq $run2["projection_metadata_overlap_sha256"])
    rule_evaluation_trace_preview_overlap_sha256_match = ($run1["rule_evaluation_trace_preview_overlap_sha256"] -ceq $run2["rule_evaluation_trace_preview_overlap_sha256"])
    projection_metadata_key_order_consistent = (($run1["projection_metadata_key_order_actual"] -join "|") -ceq ($run2["projection_metadata_key_order_actual"] -join "|"))
  }
  overall_pass = (
    $run1["overlap_byte_identical"] -and
    $run2["overlap_byte_identical"] -and
    $run1["projection_metadata_key_order_match"] -and
    $run2["projection_metadata_key_order_match"] -and
    (($run1["projection_metadata_overlap_sha256"] -ceq $run2["projection_metadata_overlap_sha256"])) -and
    (($run1["rule_evaluation_trace_preview_overlap_sha256"] -ceq $run2["rule_evaluation_trace_preview_overlap_sha256"])) -and
    ((($run1["projection_metadata_key_order_actual"] -join "|") -ceq ($run2["projection_metadata_key_order_actual"] -join "|")))
  )
}

$reportDir = Split-Path -Parent $ReportOut
if (-not [string]::IsNullOrWhiteSpace($reportDir) -and -not (Test-Path $reportDir)) {
  New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
}

$reportJson = $report | ConvertTo-Json -Depth 20
Set-Content -Path $ReportOut -Value $reportJson -Encoding utf8NoBOM

Write-Output ("overall_pass=" + $report["overall_pass"])
Write-Output ("report_out=" + $ReportOut)
Write-Output ("run1_overlap_byte_identical=" + $run1["overlap_byte_identical"])
Write-Output ("run2_overlap_byte_identical=" + $run2["overlap_byte_identical"])
Write-Output ("run1_projection_metadata_key_order_match=" + $run1["projection_metadata_key_order_match"])
Write-Output ("run2_projection_metadata_key_order_match=" + $run2["projection_metadata_key_order_match"])
