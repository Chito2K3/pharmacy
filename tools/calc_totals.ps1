$json = Get-Content items.json | ConvertFrom-Json
$disp=0
$stor=0
$ware=0
$cons=0
$total=0
$cnt = 0
foreach ($i in $json) {
    $cnt++
    $dv = [double]($i.dispensing_inventory_qty -replace "[^0-9\.\-]","")
    $sv = [double]($i.storage_inventory_qty -replace "[^0-9\.\-]","")
    $wv = [double]($i.warehouse_inventory_qty -replace "[^0-9\.\-]","")
    $cv = [double]($i.consignment_inventory_qty -replace "[^0-9\.\-]","")
    $tv = [double]($i.total_inventory_qty -replace "[^0-9\.\-]","")
    $disp += $dv
    $stor += $sv
    $ware += $wv
    $cons += $cv
    $total += $tv
    if ($dv -ne 0 -and $cnt -le 5) {
        Write-Host "Item: $($i.description.Substring(0, [Math]::Min(40,$i.description.Length))) disp=$dv stor=$sv ware=$wv"
    }
}
Write-Host "---"
Write-Host "Items count: $cnt"
Write-Host "Total: $total"
Write-Host "Dispensing: $disp"
Write-Host "Storage: $stor"
Write-Host "Warehouse: $ware"
Write-Host "Consignment: $cons"
