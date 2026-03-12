# Audit-RepoEncoding.ps1
[CmdletBinding()]
param(
    [string]$Root = "."
)

$ErrorActionPreference = "Stop"
$extensions = @(".md", ".ini", ".vbs", ".ps1")

function Get-FileEncodingInfo {
    param([byte[]]$Bytes)

    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return "UTF-8 BOM"
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return "UTF-16 LE BOM"
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
        return "UTF-16 BE BOM"
    }
    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE -and $Bytes[2] -eq 0x00 -and $Bytes[3] -eq 0x00) {
        return "UTF-32 LE BOM"
    }
    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0x00 -and $Bytes[1] -eq 0x00 -and $Bytes[2] -eq 0xFE -and $Bytes[3] -eq 0xFF) {
        return "UTF-32 BE BOM"
    }

    try {
        $utf8Strict = [System.Text.UTF8Encoding]::new($false, $true)
        [void]$utf8Strict.GetString($Bytes)
        return "UTF-8"
    }
    catch {
        return "Windows-1252/Legacy"
    }
}

$files = Get-ChildItem -Path $Root -Recurse -File |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() }

Write-Host "Matched files: $($files.Count)"
Write-Host ""

$rows = foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    [pscustomobject]@{
        Encoding = Get-FileEncodingInfo -Bytes $bytes
        File     = $file.FullName
    }
}

$rows | Group-Object Encoding | Sort-Object Name | ForEach-Object {
    "{0,-22} {1,5}" -f $_.Name, $_.Count
}

Write-Host ""
Write-Host "Files by encoding:"
$rows | Sort-Object Encoding, File | Format-Table -AutoSize
