param(
  [string]$WorkspaceFolder
)

$dir = Join-Path $WorkspaceFolder "Main_TrioScript"
$f = Get-ChildItem $dir -Filter '*.vbs' | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $f) {
  Write-Host "No .vbs found in Main_TrioScript"
  exit 1
}

$content = Get-Content $f.FullName -Raw
Set-Clipboard -Value $content
Write-Host ("Copied to clipboard: " + $f.Name)
