param(
    [Parameter(Mandatory=$true)]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [int]$Start,

    [Parameter(Mandatory=$true)]
    [int]$End
)

if (!(Test-Path $Path)) {
    Write-Error "File not found: $Path"
    exit 1
}

$c = Get-Content $Path
$total = $c.Length

if ($Start -lt 1) { $Start = 1 }
if ($End -gt $total) { $End = $total }

for ($i = $Start; $i -le $End; $i++) {
    "{0,5}: {1}" -f $i, $c[$i-1]
}
