$ErrorActionPreference = "Stop"

$base = Join-Path $PSScriptRoot ""

$files = [ordered]@{
  STRICT_RUN1_SUCCESS = Join-Path $base "tmp_phase4_strict_run1_success.txt"
  STRICT_RUN1_AMBIG = Join-Path $base "tmp_phase4_strict_run1_ambiguity.txt"
  STRICT_RUN2_SUCCESS = Join-Path $base "tmp_phase4_strict_run2_success.txt"
  STRICT_RUN2_AMBIG = Join-Path $base "tmp_phase4_strict_run2_ambiguity.txt"
  NONSTRICT_RUN1_SUCCESS = Join-Path $base "tmp_phase4_nonstrict_run1_success.txt"
  NONSTRICT_RUN1_AMBIG = Join-Path $base "tmp_phase4_nonstrict_run1_ambiguity.txt"
  NONSTRICT_RUN2_SUCCESS = Join-Path $base "tmp_phase4_nonstrict_run2_success.txt"
  NONSTRICT_RUN2_AMBIG = Join-Path $base "tmp_phase4_nonstrict_run2_ambiguity.txt"
}

foreach($k in $files.Keys){
  if(-not (Test-Path $files[$k])){ throw "Missing artifact file: $($files[$k])" }
}

$hash = @{}
foreach($k in $files.Keys){
  $hash[$k] = (Get-FileHash -Algorithm SHA256 -Path $files[$k]).Hash
}

function Get-AmbSummary([string]$path){
  $line = (Get-Content $path | Where-Object { $_ -like 'AMB_SUMMARY=*' } | Select-Object -First 1)
  if($null -eq $line){ return "" }
  return $line.Substring('AMB_SUMMARY='.Length)
}

$strict_bundle_1 = ($hash.STRICT_RUN1_SUCCESS + $hash.STRICT_RUN1_AMBIG)
$strict_bundle_2 = ($hash.STRICT_RUN2_SUCCESS + $hash.STRICT_RUN2_AMBIG)
$nonstrict_bundle_1 = ($hash.NONSTRICT_RUN1_SUCCESS + $hash.NONSTRICT_RUN1_AMBIG)
$nonstrict_bundle_2 = ($hash.NONSTRICT_RUN2_SUCCESS + $hash.NONSTRICT_RUN2_AMBIG)

$strict_1_vs_2 = ($strict_bundle_1 -eq $strict_bundle_2)
$nonstrict_1_vs_2 = ($nonstrict_bundle_1 -eq $nonstrict_bundle_2)
$strict_vs_nonstrict_success = ($hash.STRICT_RUN1_SUCCESS -eq $hash.NONSTRICT_RUN1_SUCCESS)

$strict_amb_1_vs_2 = ($hash.STRICT_RUN1_AMBIG -eq $hash.STRICT_RUN2_AMBIG)
$nonstrict_amb_1_vs_2 = ($hash.NONSTRICT_RUN1_AMBIG -eq $hash.NONSTRICT_RUN2_AMBIG)

$strict_amb_sum_1 = Get-AmbSummary $files.STRICT_RUN1_AMBIG
$strict_amb_sum_2 = Get-AmbSummary $files.STRICT_RUN2_AMBIG
$nonstrict_amb_sum_1 = Get-AmbSummary $files.NONSTRICT_RUN1_AMBIG
$nonstrict_amb_sum_2 = Get-AmbSummary $files.NONSTRICT_RUN2_AMBIG

$strict_amb_summary_equal = ($strict_amb_sum_1 -eq $strict_amb_sum_2)
$nonstrict_amb_summary_equal = ($nonstrict_amb_sum_1 -eq $nonstrict_amb_sum_2)

$phase4_pass = (
  $strict_1_vs_2 -and
  $nonstrict_1_vs_2 -and
  $strict_vs_nonstrict_success -and
  $strict_amb_1_vs_2 -and
  $nonstrict_amb_1_vs_2 -and
  $strict_amb_summary_equal -and
  $nonstrict_amb_summary_equal
)

$reportPath = Join-Path $base "tmp_phase4_hash_report.txt"
$out = @()
$out += "STRICT_RUN1_SUCCESS_SHA256=$($hash.STRICT_RUN1_SUCCESS)"
$out += "STRICT_RUN1_AMBIG_SHA256=$($hash.STRICT_RUN1_AMBIG)"
$out += "STRICT_RUN2_SUCCESS_SHA256=$($hash.STRICT_RUN2_SUCCESS)"
$out += "STRICT_RUN2_AMBIG_SHA256=$($hash.STRICT_RUN2_AMBIG)"
$out += "NONSTRICT_RUN1_SUCCESS_SHA256=$($hash.NONSTRICT_RUN1_SUCCESS)"
$out += "NONSTRICT_RUN1_AMBIG_SHA256=$($hash.NONSTRICT_RUN1_AMBIG)"
$out += "NONSTRICT_RUN2_SUCCESS_SHA256=$($hash.NONSTRICT_RUN2_SUCCESS)"
$out += "NONSTRICT_RUN2_AMBIG_SHA256=$($hash.NONSTRICT_RUN2_AMBIG)"
$out += "STRICT_1_vs_STRICT_2_HASH_EQUAL=$strict_1_vs_2"
$out += "NONSTRICT_1_vs_NONSTRICT_2_HASH_EQUAL=$nonstrict_1_vs_2"
$out += "STRICT_1_vs_NONSTRICT_1_SUCCESS_HASH_EQUAL=$strict_vs_nonstrict_success"
$out += "STRICT_AMBIG_1_vs_2_HASH_EQUAL=$strict_amb_1_vs_2"
$out += "NONSTRICT_AMBIG_1_vs_2_HASH_EQUAL=$nonstrict_amb_1_vs_2"
$out += "STRICT_AMBIG_SUMMARY_1_vs_2_EQUAL=$strict_amb_summary_equal"
$out += "NONSTRICT_AMBIG_SUMMARY_1_vs_2_EQUAL=$nonstrict_amb_summary_equal"
$out += "PHASE4_PASS=$phase4_pass"
$out | Set-Content -Path $reportPath -Encoding ASCII

Get-Content $reportPath
