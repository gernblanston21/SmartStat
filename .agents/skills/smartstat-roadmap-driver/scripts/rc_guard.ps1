param(
  [Parameter(Mandatory=$true)][string]$Mode,
  [Parameter(Mandatory=$true)][string]$ChangeDescription,
  [Parameter(Mandatory=$false)][string]$SessionPath = "SESSION.md"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ReadFileOrThrow([string]$p) {
  if (-not (Test-Path -LiteralPath $p)) {
    throw "Missing required file: $p"
  }
  return Get-Content -LiteralPath $p -Raw
}

function NormalizeText([string]$s) {
  if ($null -eq $s) { return "" }
  # normalize: lower + collapse whitespace + strip punctuation-ish separators
  $t = $s.ToLowerInvariant()
  $t = $t -replace '[\r\n\t]+', ' '
  $t = $t -replace '\s+', ' '
  $t = $t.Trim()
  return $t
}

function ExtractBulletBlock([string]$md, [string]$heading) {
  # Extract bullets immediately following a line that starts with the heading text.
  # Stops when a blank line occurs after bullets, or when a new heading starts.
  $lines = $md -split "`r?`n"
  $found = $false
  $bullets = New-Object System.Collections.Generic.List[string]

  for ($i=0; $i -lt $lines.Count; $i++) {
    $ln = $lines[$i]

    if (-not $found) {
      if ($ln.Trim() -eq $heading) {
        $found = $true
      }
      continue
    }

    # Once found: collect "- " bullets
    $trim = $ln.Trim()

    # stop if we hit another heading
    if ($trim -match '^\s*#{1,6}\s+' -or $trim -match '^\s*##\s+' -or $trim -match '^\s*###\s+') {
      break
    }

    # stop if we hit a non-bullet after we've already captured some bullets and the line is blank
    if ($bullets.Count -gt 0 -and $trim -eq "") {
      break
    }

    if ($trim -like "- *") {
      $b = $trim.Substring(2).Trim()
      if ($b) { $bullets.Add($b) | Out-Null }
      continue
    }

    # Ignore non-bullet lines until bullets start; but once bullets started,
    # non-bullet non-blank lines end the block (fail-closed behavior).
    if ($bullets.Count -gt 0 -and $trim -ne "") {
      break
    }
  }

  return $bullets
}

function TokenizeBullet([string]$bullet) {
  # Produce a deterministic set of "match phrases" for this bullet.
  # Strategy: use the full bullet as a phrase, plus key 2+ word subphrases split on punctuation.
  $norm = NormalizeText $bullet
  $phrases = New-Object System.Collections.Generic.List[string]
  if ($norm) { $phrases.Add($norm) | Out-Null }

  # Split on common separators to create additional phrases
  $parts = $norm -split '[;:,()\[\]\.]+'
  foreach ($p in $parts) {
    $pp = ($p -replace '\s+', ' ').Trim()
    if ($pp -and $pp.Length -ge 6) {
      $phrases.Add($pp) | Out-Null
    }
  }

  # Unique, stable ordering
  return ($phrases | Select-Object -Unique)
}

function MatchesAny([string]$descNorm, [System.Collections.Generic.List[string]]$bullets) {
  foreach ($b in $bullets) {
    $phrases = TokenizeBullet $b
    foreach ($ph in $phrases) {
      if ($ph -and $descNorm -like "*$ph*") {
        return ,@($true, $b)  # matched, return the bullet that triggered
      }
    }
  }
  return ,@($false, $null)
}

$modeNorm = $Mode.Trim().ToUpperInvariant()
$descRaw = $ChangeDescription.Trim()

if ($modeNorm -ne "RC") {
  Write-Host "PASS: Mode is not RC."
  exit 0
}

# RC mode requires SESSION.md to drive allowed/prohibited lists
$session = ReadFileOrThrow $SessionPath

$allowedHeading = "Allowed changes for RC1:"
$prohibHeading  = "Prohibited changes for RC1:"

$allowed = ExtractBulletBlock $session $allowedHeading
$prohib  = ExtractBulletBlock $session $prohibHeading

if ($allowed.Count -eq 0) {
  Write-Host "FAIL: RC guard could not parse any bullets under '$allowedHeading' in $SessionPath"
  exit 2
}
if ($prohib.Count -eq 0) {
  Write-Host "FAIL: RC guard could not parse any bullets under '$prohibHeading' in $SessionPath"
  exit 2
}

$descNorm = NormalizeText $descRaw

# 1) Prohibited wins
$pm = MatchesAny $descNorm $prohib
if ($pm[0] -eq $true) {
  Write-Host "FAIL: RC mode: change matches PROHIBITED rule."
  Write-Host "MATCHED_PROHIBITED_BULLET=$($pm[1])"
  Write-Host "DESC=$descRaw"
  exit 2
}

# 2) Must match an Allowed bullet explicitly
$am = MatchesAny $descNorm $allowed
if ($am[0] -eq $true) {
  Write-Host "PASS: RC mode: change matches ALLOWED rule."
  Write-Host "MATCHED_ALLOWED_BULLET=$($am[1])"
  exit 0
}

Write-Host "FAIL: RC mode: change does not match any explicit ALLOWED bullet in SESSION.md (fail-closed)."
Write-Host "DESC=$descRaw"
Write-Host ""
Write-Host "ALLOWED_BULLETS:"
$allowed | ForEach-Object { Write-Host " - $_" }
exit 2
