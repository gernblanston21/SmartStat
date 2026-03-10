[CmdletBinding()]
param(
    [string]$EvidenceRoot = "tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence",
    [string]$SummaryOut = "",
    [switch]$AllowMissing
)

$ErrorActionPreference = "Stop"

function Get-Sha256Text {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "")
    }
    finally {
        $sha.Dispose()
    }
}

function Normalize-SnapshotText {
    param([string]$Path)
    $lines = Get-Content -LiteralPath $Path
    $normalized = $lines |
        ForEach-Object { $_.Trim().ToUpperInvariant() } |
        Sort-Object
    return ($normalized -join "`n")
}

function Ensure-ParentFolder {
    param([string]$Path)
    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

if ([string]::IsNullOrWhiteSpace($SummaryOut)) {
    $SummaryOut = Join-Path $EvidenceRoot "live_validation_summary.md"
}

$required = [ordered]@{
    gateoff_baseline_snapshot = Join-Path $EvidenceRoot "gateoff_baseline/live_gateoff_baseline_snapshot.txt"
    gateoff_successor_snapshot = Join-Path $EvidenceRoot "gateoff_successor/live_gateoff_successor_snapshot.txt"
    gateon_run1_json = Join-Path $EvidenceRoot "gateon_run1/live_gateon_ingress_run1.json"
    gateon_run2_json = Join-Path $EvidenceRoot "gateon_run2/live_gateon_ingress_run2.json"
}

$optional = [ordered]@{
    gateon_run1_log = Join-Path $EvidenceRoot "gateon_run1/live_gateon_ingress_run1.log"
    gateon_run2_log = Join-Path $EvidenceRoot "gateon_run2/live_gateon_ingress_run2.log"
}

$missingRequired = New-Object System.Collections.ArrayList
foreach ($k in $required.Keys) {
    if (-not (Test-Path -LiteralPath $required[$k])) {
        [void]$missingRequired.Add($required[$k])
    }
}

$gateOffBaselineHash = ""
$gateOffSuccessorHash = ""
$gateOffMatch = $false
$gateOffStatus = "INCOMPLETE"

if (($missingRequired -notcontains $required["gateoff_baseline_snapshot"]) -and ($missingRequired -notcontains $required["gateoff_successor_snapshot"])) {
    $baselineNormalized = Normalize-SnapshotText -Path $required["gateoff_baseline_snapshot"]
    $successorNormalized = Normalize-SnapshotText -Path $required["gateoff_successor_snapshot"]
    $gateOffBaselineHash = Get-Sha256Text -Text $baselineNormalized
    $gateOffSuccessorHash = Get-Sha256Text -Text $successorNormalized
    $gateOffMatch = ($gateOffBaselineHash -eq $gateOffSuccessorHash)
    if ($gateOffMatch) {
        $gateOffStatus = "PASS"
    }
    else {
        $gateOffStatus = "FAIL"
    }
}

$gateOnRun1Hash = ""
$gateOnRun2Hash = ""
$gateOnMatch = $false
$gateOnStatus = "INCOMPLETE"

if (($missingRequired -notcontains $required["gateon_run1_json"]) -and ($missingRequired -notcontains $required["gateon_run2_json"])) {
    $gateOnRun1Hash = (Get-FileHash -LiteralPath $required["gateon_run1_json"] -Algorithm SHA256).Hash
    $gateOnRun2Hash = (Get-FileHash -LiteralPath $required["gateon_run2_json"] -Algorithm SHA256).Hash
    $gateOnMatch = ($gateOnRun1Hash -eq $gateOnRun2Hash)
    if ($gateOnMatch) {
        $gateOnStatus = "PASS"
    }
    else {
        $gateOnStatus = "FAIL"
    }
}

$optionalRun1LogHash = ""
$optionalRun2LogHash = ""
if ((Test-Path -LiteralPath $optional["gateon_run1_log"]) -and (Test-Path -LiteralPath $optional["gateon_run2_log"])) {
    $optionalRun1LogHash = (Get-FileHash -LiteralPath $optional["gateon_run1_log"] -Algorithm SHA256).Hash
    $optionalRun2LogHash = (Get-FileHash -LiteralPath $optional["gateon_run2_log"] -Algorithm SHA256).Hash
}

$overallStatus = "PASS"
if ($gateOffStatus -eq "FAIL" -or $gateOnStatus -eq "FAIL") {
    $overallStatus = "FAIL"
}
elseif ($gateOffStatus -eq "INCOMPLETE" -or $gateOnStatus -eq "INCOMPLETE") {
    $overallStatus = "INCOMPLETE"
}

$lines = New-Object System.Collections.ArrayList
[void]$lines.Add("# Runtime Slice-1 Live Validation Summary")
[void]$lines.Add("")
[void]$lines.Add("Generated: $(Get-Date -Format s)")
[void]$lines.Add("Evidence root: $EvidenceRoot")
[void]$lines.Add("")
[void]$lines.Add("## Status")
[void]$lines.Add("")
[void]$lines.Add("- Overall: $overallStatus")
[void]$lines.Add("- Gate-OFF parity: $gateOffStatus")
[void]$lines.Add("- Gate-ON determinism: $gateOnStatus")
[void]$lines.Add("")
[void]$lines.Add("## Gate-OFF Parity")
[void]$lines.Add("")
[void]$lines.Add("- Baseline file: $($required['gateoff_baseline_snapshot'])")
[void]$lines.Add("- Successor file: $($required['gateoff_successor_snapshot'])")
[void]$lines.Add("- Baseline normalized SHA256: $gateOffBaselineHash")
[void]$lines.Add("- Successor normalized SHA256: $gateOffSuccessorHash")
[void]$lines.Add("- Hash match: $gateOffMatch")
[void]$lines.Add("")
[void]$lines.Add("## Gate-ON Determinism")
[void]$lines.Add("")
[void]$lines.Add("- Run1 JSON file: $($required['gateon_run1_json'])")
[void]$lines.Add("- Run2 JSON file: $($required['gateon_run2_json'])")
[void]$lines.Add("- Run1 SHA256: $gateOnRun1Hash")
[void]$lines.Add("- Run2 SHA256: $gateOnRun2Hash")
[void]$lines.Add("- Hash match: $gateOnMatch")
[void]$lines.Add("")
[void]$lines.Add("## Optional Log Hashes")
[void]$lines.Add("")
[void]$lines.Add("- Run1 log: $($optional['gateon_run1_log'])")
[void]$lines.Add("- Run2 log: $($optional['gateon_run2_log'])")
[void]$lines.Add("- Run1 log SHA256: $optionalRun1LogHash")
[void]$lines.Add("- Run2 log SHA256: $optionalRun2LogHash")
[void]$lines.Add("")
[void]$lines.Add("## Missing Required Inputs")
[void]$lines.Add("")
if ($missingRequired.Count -eq 0) {
    [void]$lines.Add("- None")
}
else {
    foreach ($m in $missingRequired) {
        [void]$lines.Add("- $m")
    }
}

Ensure-ParentFolder -Path $SummaryOut
$lines -join "`r`n" | Set-Content -LiteralPath $SummaryOut -Encoding UTF8

Write-Output "summary_file=$SummaryOut"
Write-Output "overall_status=$overallStatus"
Write-Output "gateoff_status=$gateOffStatus"
Write-Output "gateon_status=$gateOnStatus"

if ($overallStatus -eq "PASS") {
    exit 0
}

if ($overallStatus -eq "INCOMPLETE" -and $AllowMissing) {
    exit 0
}

if ($overallStatus -eq "FAIL") {
    exit 1
}

exit 2
