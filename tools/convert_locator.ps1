$lines = Get-Content "C:\Users\PC 127\.gemini\antigravity-ide\brain\eaac2829-85d0-462a-84d9-3736cded322b\.system_generated\steps\79\content.md"
$csvLines = $lines | Select-Object -Skip 8
$csv = $csvLines | ConvertFrom-Csv

$list = @()
foreach ($row in $csv) {
    $itemCode = ""
    $dciCode = ""
    $desc = ""
    $loc = ""

    foreach ($p in $row.psobject.properties) {
        $name = $p.Name.Trim().ToLower()
        if ($name -eq "item code") { $itemCode = $p.Value }
        elseif ($name -eq "dci code") { $dciCode = $p.Value }
        elseif ($name -eq "item description") { $desc = $p.Value }
        elseif ($name -eq "location") { $loc = $p.Value }
    }

    if ($itemCode -match "0213" -and $itemCode -match "DMR") {
        $loc = "L6"
    }

    if ($itemCode) {
        $list += [PSCustomObject]@{
            item_code   = $itemCode
            dci_code    = $dciCode
            description = $desc
            location    = $loc
        }
    }
}

$json = $list | ConvertTo-Json -Depth 4
$json | Set-Content -Path "c:\Users\PC 127\Downloads\chito\pharmacy'\IUR\locator.json" -Encoding UTF8

$jsContent = "window.LOCATOR_DATA = " + $json + ";"
$jsContent | Set-Content -Path "c:\Users\PC 127\Downloads\chito\pharmacy'\IUR\locator_data.js" -Encoding UTF8

Write-Host "Done updating offline dataset to L6."
