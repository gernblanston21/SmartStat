# Normalize-RepoEncoding.ps1
# Converts .md, .ini, .vbs, .ps1 files to UTF-8 (no BOM)
# while preserving existing line endings exactly.

[CmdletBinding()]
param(
    [string]$Root = "."
)

$ErrorActionPreference = "Stop"

$extensions = @(".md", ".ini", ".vbs", ".ps1")

function Get-FileEncodingInfo {
    param(
        [byte[]]$Bytes
    )

    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return @{
            Name   = "UTF-8 BOM"
            PreambleLength = 3
            Encoding = [System.Text.UTF8Encoding]::new($true, $true)
        }
    }

    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return @{
            Name   = "UTF-16 LE BOM"
            PreambleLength = 2
            Encoding = [System.Text.UnicodeEncoding]::new($false, $true, $true)
        }
    }

    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
        return @{
            Name   = "UTF-16 BE BOM"
            PreambleLength = 2
            Encoding = [System.Text.UnicodeEncoding]::new($true, $true, $true)
        }
    }

    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE -and $Bytes[2] -eq 0x00 -and $Bytes[3] -eq 0x00) {
        return @{
            Name   = "UTF-32 LE BOM"
            PreambleLength = 4
            Encoding = [System.Text.UTF32Encoding]::new($false, $true, $true)
        }
    }

    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0x00 -and $Bytes[1] -eq 0x00 -and $Bytes[2] -eq 0xFE -and $Bytes[3] -eq 0xFF) {
        return @{
            Name   = "UTF-32 BE BOM"
            PreambleLength = 4
            Encoding = [System.Text.UTF32Encoding]::new($true, $true, $true)
        }
    }

    # No BOM: try strict UTF-8 first.
    try {
        $utf8Strict = [System.Text.UTF8Encoding]::new($false, $true)
        [void]$utf8Strict.GetString($Bytes)
        return @{
            Name   = "UTF-8"
            PreambleLength = 0
            Encoding = $utf8Strict
        }
    }
    catch {
        # Fall back to Windows-1252 for legacy Windows-authored files.
        return @{
            Name   = "Windows-1252"
            PreambleLength = 0
            Encoding = [System.Text.Encoding]::GetEncoding(1252)
        }
    }
}

function Get-LineEndingStyle {
    param(
        [byte[]]$Bytes
    )

    $hasCRLF = $false
    $hasLFOnly = $false
    $hasCROnly = $false

    for ($i = 0; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i] -eq 0x0D) {
            if ($i + 1 -lt $Bytes.Length -and $Bytes[$i + 1] -eq 0x0A) {
                $hasCRLF = $true
                $i++
            }
            else {
                $hasCROnly = $true
            }
        }
        elseif ($Bytes[$i] -eq 0x0A) {
            $hasLFOnly = $true
        }
    }

    if ($hasCRLF -and -not $hasLFOnly -and -not $hasCROnly) { return "CRLF" }
    if ($hasLFOnly -and -not $hasCRLF -and -not $hasCROnly) { return "LF" }
    if ($hasCROnly -and -not $hasCRLF -and -not $hasLFOnly) { return "CR" }
    if (-not $hasCRLF -and -not $hasLFOnly -and -not $hasCROnly) { return "None" }

    return "Mixed"
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

$files = Get-ChildItem -Path $Root -Recurse -File |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() }

if (-not $files) {
    Write-Host "No matching files found."
    exit 0
}

$converted = 0

foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $encodingInfo = Get-FileEncodingInfo -Bytes $bytes
    $lineEnding = Get-LineEndingStyle -Bytes $bytes

    try {
        $text = $encodingInfo.Encoding.GetString($bytes, $encodingInfo.PreambleLength, $bytes.Length - $encodingInfo.PreambleLength)
    }
    catch {
        Write-Warning "Skipping unreadable file: $($file.FullName)"
        continue
    }

    # Re-encode as UTF-8 without BOM.
    $newBytes = $utf8NoBom.GetBytes($text)

    # Only write if bytes actually change.
    $sameLength = ($bytes.Length -eq $newBytes.Length)
    $sameBytes = $sameLength
    if ($sameBytes) {
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -ne $newBytes[$i]) {
                $sameBytes = $false
                break
            }
        }
    }

    if (-not $sameBytes) {
        [System.IO.File]::WriteAllBytes($file.FullName, $newBytes)
        $converted++
        Write-Host ("Converted: {0} | {1} -> UTF-8 | Line endings: {2}" -f $file.FullName, $encodingInfo.Name, $lineEnding)
    }
    else {
        Write-Host ("OK:        {0} | {1} | Line endings: {2}" -f $file.FullName, $encodingInfo.Name, $lineEnding)
    }
}

Write-Host ""
Write-Host "Done. Files rewritten: $converted"
