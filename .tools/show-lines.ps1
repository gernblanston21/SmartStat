param(
    [string]$Path,
    [int]$Start,
    [int]$End
)

$c = Get-Content $Path
for ($i = $Start; $i -le $End; $i++) {
    if ($i -lt $c.Length) {
        "{0,5}: {1}" -f ($i + 1), $c[$i]
    }
}
