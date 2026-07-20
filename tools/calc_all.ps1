$json = Get-Content items.json | ConvertFrom-Json
$sums = @{}
foreach ($i in $json) {
    foreach ($p in $i.psobject.properties) {
        if ($p.Value -match "^[\d\.]+$") {
            if (-not $sums.ContainsKey($p.Name)) { $sums[$p.Name] = 0 }
            $sums[$p.Name] += [double]$p.Value
        }
    }
}
$sums.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    Write-Host "$($_.Name): $($_.Value)"
}
