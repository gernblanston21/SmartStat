param(
  [Parameter(Mandatory = $false)][string]$Path = "",
  [Parameter(Mandatory = $false)][string]$RunLabel = "",
  [Parameter(Mandatory = $false)][string]$CaseLabel = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$commonPath = Join-Path $PSScriptRoot "common_plan_capture_validator.ps1"
if (-not (Test-Path -LiteralPath $commonPath)) {
  throw "Missing helper script: $commonPath"
}
. $commonPath

function Test-RequiredKeys {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)]$ObjectValue,
    [Parameter(Mandatory = $true)][string[]]$RequiredKeys,
    [Parameter(Mandatory = $true)][string]$Path
  )

  foreach ($requiredKey in $RequiredKeys) {
    if (-not $ObjectValue.ContainsKey($requiredKey)) {
      Add-Issue -Issues $Issues -Code "MISSING_REQUIRED_KEY" -Path $Path -Message "Missing required key '$requiredKey'."
    }
  }
}

function Test-AllowedKeys {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)]$ObjectValue,
    [Parameter(Mandatory = $true)][string[]]$AllowedKeys,
    [Parameter(Mandatory = $true)][string]$Path
  )

  $allowedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
  foreach ($k in $AllowedKeys) {
    $allowedSet.Add($k) | Out-Null
  }

  foreach ($key in @($ObjectValue.Keys | ForEach-Object { [string]$_ })) {
    if (-not $allowedSet.Contains($key)) {
      Add-Issue -Issues $Issues -Code "UNKNOWN_KEY" -Path $Path -Message "Unknown key '$key' is not allowed."
    }
  }
}

function Test-NonEmptyString {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)]$ObjectValue,
    [Parameter(Mandatory = $true)][string]$Key,
    [Parameter(Mandatory = $true)][string]$Path
  )

  if (-not $ObjectValue.ContainsKey($Key)) {
    return
  }

  $value = $ObjectValue[$Key]
  if (-not ($value -is [string]) -or [string]::IsNullOrWhiteSpace([string]$value)) {
    Add-Issue -Issues $Issues -Code "INVALID_STRING" -Path "$Path.$Key" -Message "Key '$Key' must be a non-empty string."
  }
}

function Test-EnumValue {
  param(
    [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Issues,
    [Parameter(Mandatory = $true)]$ObjectValue,
    [Parameter(Mandatory = $true)][string]$Key,
    [Parameter(Mandatory = $true)][string[]]$AllowedValues,
    [Parameter(Mandatory = $true)][string]$Path
  )

  if (-not $ObjectValue.ContainsKey($Key)) {
    return
  }

  $value = $ObjectValue[$Key]
  if (-not ($value -is [string])) {
    Add-Issue -Issues $Issues -Code "INVALID_ENUM_VALUE" -Path "$Path.$Key" -Message "Key '$Key' must be a string."
    return
  }

  if (-not ($AllowedValues -contains [string]$value)) {
    Add-Issue -Issues $Issues -Code "INVALID_ENUM_VALUE" -Path "$Path.$Key" -Message "Key '$Key' value '$value' is not allowed."
  }
}

$validatorId = "validate_plan_capture_contract"
$contractVersion = "wp17.plan_capture.v1"

$requiredTopKeys = @("contract_version", "artifact_type", "mode", "source", "determinism", "slot_sequence", "terminal", "refusal", "deferred_boundaries")
$requiredSourceKeys = @("query_text", "normalized_query_text", "semantic_record_id", "semantic_record_type", "league")
$requiredDeterminismKeys = @("ordering_contract", "canonicalization_version", "input_fingerprint_sha256")
$requiredSlotKeys = @("order", "slot_class", "token", "token_kind", "parameters", "source_dictionary", "candidate_status")
$requiredTerminalKeys = @("slot_class", "token")
$requiredRefusalKeys = @("stage", "code", "message", "blocking_slot_order")

$allowedSlotClasses = @("family", "operator", "entity", "scope", "filter", "terminal_measure", "terminal_attribute", "formatter")
$allowedTokenKinds = @("canonical", "alias_normalized", "literal_parameter")
$allowedSourceDictionaries = @("query_skeletons", "operator_grammar", "entity_dictionary", "filter_grammar_dictionary", "measure_dictionary", "attribute_dictionary", "formatter_dictionary", "derived")
$allowedCandidateStatuses = @("resolved", "inferred")
$allowedArtifactTypes = @("captured_plan", "capture_refusal")
$allowedRefusalCodes = @("AMBIGUOUS_SLOT", "MISSING_TERMINAL", "UNSUPPORTED_SLOT_CLASS", "ORDER_VIOLATION", "UNSUPPORTED_OPERATOR_FORM", "UNSUPPORTED_STATE")
$requiredDeferredBoundaries = @("runtime apply behavior", "planner execution", "runtime bridge behavior")

$repoRoot = Find-RepoRoot -StartDir $PSScriptRoot
$defaultTarget = Join-Path "tests" "wp-17\fixtures\good\good_captured_plan_basic.json"
$resolvedPath = Resolve-InputPath -RepoRoot $repoRoot -Path $Path -DefaultFileName $defaultTarget
$runLabelFinal = New-RunLabel -RunLabel $RunLabel
$caseTokenSource = if ([string]::IsNullOrWhiteSpace($CaseLabel)) { [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath) } else { $CaseLabel }
$caseToken = Get-SafeToken -Value $caseTokenSource -Fallback "plan_capture"

$artifactDir = Join-Path (Join-Path $PSScriptRoot "artifacts") $runLabelFinal
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null
$jsonPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.json"
$txtPath = Join-Path $artifactDir "$validatorId.$caseToken.summary.txt"

$issues = New-IssueList
$parseable = $true
$payload = $null
$canonicalJson = ""
$canonicalSha = ""

if (-not (Test-Path -LiteralPath $resolvedPath)) {
  Add-Issue -Issues $issues -Code "FILE_NOT_FOUND" -Path "" -Message "JSON path not found: $resolvedPath"
  $parseable = $false
}
else {
  try {
    $raw = [System.IO.File]::ReadAllText($resolvedPath)
    $convertFromJsonCmd = Get-Command ConvertFrom-Json
    if (-not $convertFromJsonCmd.Parameters.ContainsKey("AsHashtable")) {
      throw "ConvertFrom-Json -AsHashtable is required. Use PowerShell 7+."
    }
    $payload = ConvertFrom-Json -InputObject $raw -AsHashtable -Depth 100
    if ($null -eq $payload -or -not ($payload -is [System.Collections.IDictionary])) {
      Add-Issue -Issues $issues -Code "INVALID_TOP_LEVEL" -Path "" -Message "Top-level JSON value must be an object."
    }
  }
  catch {
    Add-Issue -Issues $issues -Code "JSON_PARSE_ERROR" -Path "" -Message ("Failed to parse JSON: " + $_.Exception.Message)
    $parseable = $false
  }
}

if ($parseable -and $payload -is [System.Collections.IDictionary]) {
  Test-RequiredKeys -Issues $issues -ObjectValue $payload -RequiredKeys $requiredTopKeys -Path "$"
  Test-AllowedKeys -Issues $issues -ObjectValue $payload -AllowedKeys $requiredTopKeys -Path "$"

  if ($payload.ContainsKey("contract_version")) {
    if ([string]$payload["contract_version"] -ne $contractVersion) {
      Add-Issue -Issues $issues -Code "INVALID_CONTRACT_VERSION" -Path "$.contract_version" -Message "contract_version must be '$contractVersion'."
    }
  }

  Test-EnumValue -Issues $issues -ObjectValue $payload -Key "artifact_type" -AllowedValues $allowedArtifactTypes -Path "$"
  if ($payload.ContainsKey("mode")) {
    if ([string]$payload["mode"] -ne "read_only_contract") {
      Add-Issue -Issues $issues -Code "INVALID_MODE" -Path "$.mode" -Message "mode must be 'read_only_contract'."
    }
  }

  if ($payload.ContainsKey("source")) {
    $source = $payload["source"]
    if (-not ($source -is [System.Collections.IDictionary])) {
      Add-Issue -Issues $issues -Code "INVALID_SOURCE_OBJECT" -Path "$.source" -Message "source must be an object."
    }
    else {
      Test-RequiredKeys -Issues $issues -ObjectValue $source -RequiredKeys $requiredSourceKeys -Path "$.source"
      Test-AllowedKeys -Issues $issues -ObjectValue $source -AllowedKeys $requiredSourceKeys -Path "$.source"
      Test-NonEmptyString -Issues $issues -ObjectValue $source -Key "query_text" -Path "$.source"
      Test-NonEmptyString -Issues $issues -ObjectValue $source -Key "normalized_query_text" -Path "$.source"
      Test-NonEmptyString -Issues $issues -ObjectValue $source -Key "semantic_record_id" -Path "$.source"
      Test-NonEmptyString -Issues $issues -ObjectValue $source -Key "semantic_record_type" -Path "$.source"
      if ($source.ContainsKey("league")) {
        $league = $source["league"]
        if ($null -ne $league -and -not ($league -is [string])) {
          Add-Issue -Issues $issues -Code "INVALID_LEAGUE_TYPE" -Path "$.source.league" -Message "league must be string or null."
        }
      }
    }
  }

  if ($payload.ContainsKey("determinism")) {
    $determinism = $payload["determinism"]
    if (-not ($determinism -is [System.Collections.IDictionary])) {
      Add-Issue -Issues $issues -Code "INVALID_DETERMINISM_OBJECT" -Path "$.determinism" -Message "determinism must be an object."
    }
    else {
      Test-RequiredKeys -Issues $issues -ObjectValue $determinism -RequiredKeys $requiredDeterminismKeys -Path "$.determinism"
      Test-AllowedKeys -Issues $issues -ObjectValue $determinism -AllowedKeys $requiredDeterminismKeys -Path "$.determinism"
      if ($determinism.ContainsKey("ordering_contract") -and [string]$determinism["ordering_contract"] -ne "slot_sequence.order_asc.v1") {
        Add-Issue -Issues $issues -Code "INVALID_ORDERING_CONTRACT" -Path "$.determinism.ordering_contract" -Message "ordering_contract must be 'slot_sequence.order_asc.v1'."
      }
      if ($determinism.ContainsKey("canonicalization_version") -and [string]$determinism["canonicalization_version"] -ne "wp17.canonical_json.v1") {
        Add-Issue -Issues $issues -Code "INVALID_CANONICALIZATION_VERSION" -Path "$.determinism.canonicalization_version" -Message "canonicalization_version must be 'wp17.canonical_json.v1'."
      }
      if ($determinism.ContainsKey("input_fingerprint_sha256")) {
        $fp = [string]$determinism["input_fingerprint_sha256"]
        if (-not [System.Text.RegularExpressions.Regex]::IsMatch($fp, "^[A-Fa-f0-9]{64}$")) {
          Add-Issue -Issues $issues -Code "INVALID_INPUT_FINGERPRINT" -Path "$.determinism.input_fingerprint_sha256" -Message "input_fingerprint_sha256 must be 64 hex characters."
        }
      }
    }
  }

  $slotSequence = @()
  if ($payload.ContainsKey("slot_sequence")) {
    $slotRaw = $payload["slot_sequence"]
    if ($slotRaw -isnot [System.Collections.IEnumerable] -or $slotRaw -is [string]) {
      Add-Issue -Issues $issues -Code "INVALID_SLOT_SEQUENCE" -Path "$.slot_sequence" -Message "slot_sequence must be an array."
    }
    else {
      $slotSequence = @($slotRaw)
      for ($slotIndex = 0; $slotIndex -lt $slotSequence.Count; $slotIndex++) {
        $slotPath = "$.slot_sequence[$slotIndex]"
        $slot = $slotSequence[$slotIndex]
        if (-not ($slot -is [System.Collections.IDictionary])) {
          Add-Issue -Issues $issues -Code "INVALID_SLOT_OBJECT" -Path $slotPath -Message "Each slot entry must be an object."
          continue
        }

        Test-RequiredKeys -Issues $issues -ObjectValue $slot -RequiredKeys $requiredSlotKeys -Path $slotPath
        Test-AllowedKeys -Issues $issues -ObjectValue $slot -AllowedKeys $requiredSlotKeys -Path $slotPath

        if ($slot.ContainsKey("order")) {
          $orderValue = $slot["order"]
          if ($orderValue -isnot [int] -and $orderValue -isnot [long]) {
            Add-Issue -Issues $issues -Code "INVALID_SLOT_ORDER_TYPE" -Path "$slotPath.order" -Message "order must be an integer."
          }
          elseif ([int]$orderValue -lt 1) {
            Add-Issue -Issues $issues -Code "INVALID_SLOT_ORDER_VALUE" -Path "$slotPath.order" -Message "order must be >= 1."
          }
        }

        Test-EnumValue -Issues $issues -ObjectValue $slot -Key "slot_class" -AllowedValues $allowedSlotClasses -Path $slotPath
        Test-NonEmptyString -Issues $issues -ObjectValue $slot -Key "token" -Path $slotPath
        Test-EnumValue -Issues $issues -ObjectValue $slot -Key "token_kind" -AllowedValues $allowedTokenKinds -Path $slotPath
        Test-EnumValue -Issues $issues -ObjectValue $slot -Key "source_dictionary" -AllowedValues $allowedSourceDictionaries -Path $slotPath
        Test-EnumValue -Issues $issues -ObjectValue $slot -Key "candidate_status" -AllowedValues $allowedCandidateStatuses -Path $slotPath

        if ($slot.ContainsKey("parameters")) {
          $parameters = $slot["parameters"]
          if ($parameters -isnot [System.Collections.IEnumerable] -or $parameters -is [string]) {
            Add-Issue -Issues $issues -Code "INVALID_PARAMETERS_TYPE" -Path "$slotPath.parameters" -Message "parameters must be an array of strings."
          }
          else {
            $paramIdx = 0
            foreach ($parameter in @($parameters)) {
              if ($parameter -isnot [string]) {
                Add-Issue -Issues $issues -Code "INVALID_PARAMETER_VALUE" -Path "$slotPath.parameters[$paramIdx]" -Message "parameters entries must be strings."
              }
              $paramIdx++
            }
          }
        }
      }
    }
  }

  if ($payload.ContainsKey("deferred_boundaries")) {
    $deferred = $payload["deferred_boundaries"]
    if ($deferred -isnot [System.Collections.IEnumerable] -or $deferred -is [string]) {
      Add-Issue -Issues $issues -Code "INVALID_DEFERRED_BOUNDARIES" -Path "$.deferred_boundaries" -Message "deferred_boundaries must be an array of strings."
    }
    else {
      $deferredValues = @($deferred | ForEach-Object { [string]$_ })
      if ($deferredValues.Count -lt 1) {
        Add-Issue -Issues $issues -Code "EMPTY_DEFERRED_BOUNDARIES" -Path "$.deferred_boundaries" -Message "deferred_boundaries must contain at least one value."
      }

      foreach ($requiredBoundary in $requiredDeferredBoundaries) {
        if (-not ($deferredValues -contains $requiredBoundary)) {
          Add-Issue -Issues $issues -Code "MISSING_DEFERRED_BOUNDARY" -Path "$.deferred_boundaries" -Message "Missing required deferred boundary '$requiredBoundary'."
        }
      }
    }
  }

  $artifactType = if ($payload.ContainsKey("artifact_type")) { [string]$payload["artifact_type"] } else { "" }

  $orderValues = @()
  foreach ($slot in $slotSequence) {
    if ($slot -is [System.Collections.IDictionary] -and $slot.ContainsKey("order") -and ($slot["order"] -is [int] -or $slot["order"] -is [long])) {
      $orderValues += [int]$slot["order"]
    }
  }

  if ($orderValues.Count -gt 0) {
    $distinctOrders = @($orderValues | Sort-Object -Unique)
    if ($distinctOrders.Count -ne $orderValues.Count) {
      Add-Issue -Issues $issues -Code "DUPLICATE_SLOT_ORDER" -Path "$.slot_sequence" -Message "slot_sequence contains duplicate order values."
    }

    for ($expected = 1; $expected -le $distinctOrders.Count; $expected++) {
      if ($distinctOrders[$expected - 1] -ne $expected) {
        Add-Issue -Issues $issues -Code "NON_CONTIGUOUS_SLOT_ORDER" -Path "$.slot_sequence" -Message "slot_sequence.order must be contiguous starting at 1."
        break
      }
    }
  }

  if ($artifactType -eq "captured_plan") {
    if ($slotSequence.Count -lt 1) {
      Add-Issue -Issues $issues -Code "EMPTY_CAPTURED_SLOT_SEQUENCE" -Path "$.slot_sequence" -Message "captured_plan must include at least one slot."
    }

    if ($payload.ContainsKey("refusal") -and $null -ne $payload["refusal"]) {
      Add-Issue -Issues $issues -Code "CAPTURED_PLAN_REFUSAL_NOT_NULL" -Path "$.refusal" -Message "captured_plan requires refusal=null."
    }

    $terminal = if ($payload.ContainsKey("terminal")) { $payload["terminal"] } else { $null }
    if ($null -eq $terminal -or -not ($terminal -is [System.Collections.IDictionary])) {
      Add-Issue -Issues $issues -Code "MISSING_TERMINAL" -Path "$.terminal" -Message "captured_plan requires terminal object."
    }
    else {
      Test-RequiredKeys -Issues $issues -ObjectValue $terminal -RequiredKeys $requiredTerminalKeys -Path "$.terminal"
      Test-AllowedKeys -Issues $issues -ObjectValue $terminal -AllowedKeys $requiredTerminalKeys -Path "$.terminal"
      Test-EnumValue -Issues $issues -ObjectValue $terminal -Key "slot_class" -AllowedValues @("terminal_measure", "terminal_attribute") -Path "$.terminal"
      Test-NonEmptyString -Issues $issues -ObjectValue $terminal -Key "token" -Path "$.terminal"
    }

    $slotsSorted = @(
      $slotSequence |
      Where-Object { $_ -is [System.Collections.IDictionary] -and $_.ContainsKey("order") -and ($_.ContainsKey("slot_class")) } |
      Sort-Object { [int]$_.order }
    )

    if ($slotsSorted.Count -gt 0) {
      $firstClass = [string]$slotsSorted[0]["slot_class"]
      if ($firstClass -ne "family" -and $firstClass -ne "operator") {
        Add-Issue -Issues $issues -Code "INVALID_FIRST_SLOT" -Path "$.slot_sequence[0].slot_class" -Message "First slot must be 'family' or 'operator'."
      }
    }

    $entityCount = @($slotsSorted | Where-Object { [string]$_["slot_class"] -eq "entity" }).Count
    if ($entityCount -lt 1) {
      Add-Issue -Issues $issues -Code "MISSING_ENTITY_SLOT" -Path "$.slot_sequence" -Message "captured_plan requires at least one entity slot."
    }

    $familyCount = @($slotsSorted | Where-Object { [string]$_["slot_class"] -eq "family" }).Count
    $operatorCount = @($slotsSorted | Where-Object { [string]$_["slot_class"] -eq "operator" }).Count
    if ($familyCount -gt 1) {
      Add-Issue -Issues $issues -Code "DUPLICATE_FAMILY_SLOT" -Path "$.slot_sequence" -Message "At most one family slot is allowed."
    }
    if ($operatorCount -gt 1) {
      Add-Issue -Issues $issues -Code "DUPLICATE_OPERATOR_SLOT" -Path "$.slot_sequence" -Message "At most one operator slot is allowed."
    }
    if ($familyCount -gt 0 -and $operatorCount -gt 0) {
      Add-Issue -Issues $issues -Code "UNSUPPORTED_FAMILY_OPERATOR_COMBINATION" -Path "$.slot_sequence" -Message "Simultaneous family and operator slots are unsupported in WP-17 capture contract."
    }

    $terminalSlots = @($slotsSorted | Where-Object { [string]$_["slot_class"] -eq "terminal_measure" -or [string]$_["slot_class"] -eq "terminal_attribute" })
    if ($terminalSlots.Count -ne 1) {
      Add-Issue -Issues $issues -Code "INVALID_TERMINAL_SLOT_COUNT" -Path "$.slot_sequence" -Message "captured_plan must contain exactly one terminal slot."
    }
    else {
      $terminalSlot = $terminalSlots[0]
      if ($terminal -is [System.Collections.IDictionary]) {
        if ([string]$terminalSlot["slot_class"] -ne [string]$terminal["slot_class"] -or [string]$terminalSlot["token"] -ne [string]$terminal["token"]) {
          Add-Issue -Issues $issues -Code "TERMINAL_MISMATCH" -Path "$.terminal" -Message "terminal object must match terminal slot in slot_sequence."
        }
      }

      $terminalOrder = [int]$terminalSlot["order"]
      foreach ($slot in $slotsSorted) {
        $order = [int]$slot["order"]
        $slotClass = [string]$slot["slot_class"]
        if ($order -gt $terminalOrder -and $slotClass -ne "formatter") {
          Add-Issue -Issues $issues -Code "SLOT_AFTER_TERMINAL" -Path "$.slot_sequence" -Message "Only formatter slot may appear after terminal slot."
          break
        }
      }
    }

    $formatterSlots = @($slotsSorted | Where-Object { [string]$_["slot_class"] -eq "formatter" })
    if ($formatterSlots.Count -gt 1) {
      Add-Issue -Issues $issues -Code "MULTIPLE_FORMATTERS" -Path "$.slot_sequence" -Message "At most one formatter slot is allowed."
    }
    elseif ($formatterSlots.Count -eq 1 -and $slotsSorted.Count -gt 0) {
      $lastOrder = [int]$slotsSorted[$slotsSorted.Count - 1]["order"]
      if ([int]$formatterSlots[0]["order"] -ne $lastOrder) {
        Add-Issue -Issues $issues -Code "FORMATTER_NOT_LAST" -Path "$.slot_sequence" -Message "formatter slot must be last."
      }
    }
  }
  elseif ($artifactType -eq "capture_refusal") {
    if ($payload.ContainsKey("terminal") -and $null -ne $payload["terminal"]) {
      Add-Issue -Issues $issues -Code "REFUSAL_TERMINAL_NOT_NULL" -Path "$.terminal" -Message "capture_refusal requires terminal=null."
    }

    $refusal = if ($payload.ContainsKey("refusal")) { $payload["refusal"] } else { $null }
    if ($null -eq $refusal -or -not ($refusal -is [System.Collections.IDictionary])) {
      Add-Issue -Issues $issues -Code "MISSING_REFUSAL_OBJECT" -Path "$.refusal" -Message "capture_refusal requires refusal object."
    }
    else {
      Test-RequiredKeys -Issues $issues -ObjectValue $refusal -RequiredKeys $requiredRefusalKeys -Path "$.refusal"
      Test-AllowedKeys -Issues $issues -ObjectValue $refusal -AllowedKeys $requiredRefusalKeys -Path "$.refusal"
      if ($refusal.ContainsKey("stage") -and [string]$refusal["stage"] -ne "capture") {
        Add-Issue -Issues $issues -Code "INVALID_REFUSAL_STAGE" -Path "$.refusal.stage" -Message "refusal.stage must be 'capture'."
      }
      Test-EnumValue -Issues $issues -ObjectValue $refusal -Key "code" -AllowedValues $allowedRefusalCodes -Path "$.refusal"
      Test-NonEmptyString -Issues $issues -ObjectValue $refusal -Key "message" -Path "$.refusal"
      if ($refusal.ContainsKey("blocking_slot_order")) {
        $blocking = $refusal["blocking_slot_order"]
        if ($null -ne $blocking -and $blocking -isnot [int] -and $blocking -isnot [long]) {
          Add-Issue -Issues $issues -Code "INVALID_BLOCKING_SLOT_ORDER" -Path "$.refusal.blocking_slot_order" -Message "blocking_slot_order must be integer or null."
        }
        elseif ($null -ne $blocking) {
          $slotCount = $slotSequence.Count
          if ([int]$blocking -lt 1 -or ($slotCount -gt 0 -and [int]$blocking -gt $slotCount)) {
            Add-Issue -Issues $issues -Code "BLOCKING_SLOT_ORDER_OUT_OF_RANGE" -Path "$.refusal.blocking_slot_order" -Message "blocking_slot_order is out of slot_sequence range."
          }
        }
      }
    }
  }

  try {
    $canonicalJson = Get-CanonicalJson -Value $payload
    $canonicalSha = Get-Sha256Hex -Text $canonicalJson
  }
  catch {
    Add-Issue -Issues $issues -Code "CANONICALIZATION_FAILED" -Path "" -Message ("Canonicalization failed: " + $_.Exception.Message)
  }
}

$issueArray = @($issues.ToArray())
$issueCount = $issueArray.Count
$resultPass = ($issueCount -eq 0)

$summary = [ordered]@{
  validator = $validatorId
  contract_version = $contractVersion
  run_label = $runLabelFinal
  case_label = $caseToken
  target_path = $resolvedPath
  parseable = $parseable
  issue_count = $issueCount
  canonical_sha256 = $canonicalSha
  result = if ($resultPass) { "PASS" } else { "FAIL" }
  issues = $issueArray
}

$summary | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$txtLines = [System.Collections.Generic.List[string]]::new()
$txtLines.Add("VALIDATOR=$validatorId") | Out-Null
$txtLines.Add("CONTRACT_VERSION=$contractVersion") | Out-Null
$txtLines.Add("RUN_LABEL=$runLabelFinal") | Out-Null
$txtLines.Add("CASE_LABEL=$caseToken") | Out-Null
$txtLines.Add("TARGET_PATH=$resolvedPath") | Out-Null
$txtLines.Add("PARSEABLE=$parseable") | Out-Null
$txtLines.Add("ISSUE_COUNT=$issueCount") | Out-Null
$txtLines.Add("CANONICAL_SHA256=$canonicalSha") | Out-Null
$txtLines.Add("RESULT=$($summary.result)") | Out-Null

foreach ($issue in $issueArray) {
  $txtLines.Add("ISSUE CODE=$($issue.code) PATH=$($issue.path) MSG=$($issue.message)") | Out-Null
}

$txtLines | Set-Content -LiteralPath $txtPath -Encoding ASCII

Write-Host "VALIDATOR=$validatorId"
Write-Host "CONTRACT_VERSION=$contractVersion"
Write-Host "RUN_LABEL=$runLabelFinal"
Write-Host "CASE_LABEL=$caseToken"
Write-Host "TARGET_PATH=$resolvedPath"
Write-Host "PARSEABLE=$parseable"
Write-Host "ISSUE_COUNT=$issueCount"
Write-Host "CANONICAL_SHA256=$canonicalSha"
Write-Host "RESULT=$($summary.result)"
Write-Host "JSON_SUMMARY=$jsonPath"
Write-Host "TXT_SUMMARY=$txtPath"

if ($resultPass) {
  exit 0
}

exit 2
