$json = Get-Content items.json -Raw | ConvertFrom-Json
$disp = 0
$stor = 0
$ware = 0
$cons = 0
$endingTotal = 0
$endingEpaTotal = 0
$pendingPoTotal = 0
$avgMonthlyTotal = 0
$cnt = 0

foreach ($i in $json) {
    $cnt++
    $dv = [double]($i.dispensing_inventory_qty -replace "[^0-9\.\-]","")
    $sv = [double]($i.storage_inventory_qty -replace "[^0-9\.\-]","")
    $wv = [double]($i.warehouse_inventory_qty -replace "[^0-9\.\-]","")
    $cv = [double]($i.consignment_inventory_qty -replace "[^0-9\.\-]","")
    
    # Ending Inventory Qty
    $ev = 0
    if ($null -ne $i.ending_inventory_qty) { $ev = [double]($i.ending_inventory_qty -replace "[^0-9\.\-]","") }
    
    # Ending Inventory w/ EPA
    $epa = 0
    if ($null -ne $i.ending_inv_with_epa_qty -and [double]($i.ending_inv_with_epa_qty -replace "[^0-9\.\-]","") -gt 0) {
        $epa = [double]($i.ending_inv_with_epa_qty -replace "[^0-9\.\-]","")
    } elseif ($null -ne $i.ending_with_epa_qty) {
        $epa = [double]($i.ending_with_epa_qty -replace "[^0-9\.\-]","")
    }

    # Pending PO / CO Qty
    $po = 0
    if ($null -ne $i.pending_po_co_qty) { $po = [double]($i.pending_po_co_qty -replace "[^0-9\.\-]","") }

    # Avg Monthly Consumption / Demand
    $avg = 0
    if ($null -ne $i.avg_monthly_normalized_demand) { $avg = [double]($i.avg_monthly_normalized_demand -replace "[^0-9\.\-]","") }
    elseif ($null -ne $i.avg_monthly_consumption) { $avg = [double]($i.avg_monthly_consumption -replace "[^0-9\.\-]","") }

    $disp += $dv
    $stor += $sv
    $ware += $wv
    $cons += $cv
    $endingTotal += $ev
    $endingEpaTotal += $epa
    $pendingPoTotal += $po
    $avgMonthlyTotal += $avg
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  INVENTORY UTILIZATION SUMMARY REPORT" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ("  Total Items Count               : {0:N0}" -f $cnt)
Write-Host ("  Total Ending Inventory Qty      : {0:N0}" -f $endingTotal)
Write-Host ("  Total Ending Inventory (w/ EPA) : {0:N0}" -f $endingEpaTotal)
Write-Host ("  Total Pending PO / Call-Off Qty : {0:N0}" -f $pendingPoTotal)
Write-Host ("  Total Monthly Demand (Normalized): {0:N2}" -f $avgMonthlyTotal)
Write-Host "------------------------------------------"
if (($disp + $stor + $ware + $cons) -eq 0 -and $endingTotal -gt 0) {
    Write-Host "  Sub-location Stock Breakdown    : [FALLBACK MODE] Unassigned location sub-totals" -ForegroundColor Yellow
    Write-Host ("  Primary Ending Stock Total      : {0:N0} (Active in Sheet)" -f $endingTotal) -ForegroundColor Green
} else {
    Write-Host ("  Dispensing Stock Qty            : {0:N0}" -f $disp)
    Write-Host ("  Storage Stock Qty               : {0:N0}" -f $stor)
    Write-Host ("  Warehouse Stock Qty             : {0:N0}" -f $ware)
    Write-Host ("  Consignment Stock Qty           : {0:N0}" -f $cons)
}
Write-Host "=========================================="

