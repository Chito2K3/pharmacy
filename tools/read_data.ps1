[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem")
$zip = [System.IO.Compression.ZipFile]::OpenRead("MASTERLIST_2026_V30.xlsx")

function Get-EntryXML($zip, $name) {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $content = $reader.ReadToEnd()
    $reader.Close(); $stream.Close()
    return [xml]$content
}

$ssXml = Get-EntryXML $zip "xl/sharedStrings.xml"
$sharedStrings = @()
if ($ssXml) {
    $sharedStrings = $ssXml.sst.si | ForEach-Object {
        if ($_.t) { $_.t }
        elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' }
        else { '' }
    }
}

$sheetXml = Get-EntryXML $zip "xl/worksheets/sheet3.xml"
$rows = $sheetXml.worksheet.sheetData.row

function Get-CellValue($cell) {
    $v = $cell.v
    if ($null -eq $v) { return "" }
    if ($cell.t -eq "s") { return $sharedStrings[[int]$v] }
    return $v
}

function Col-Index($ref) {
    $colStr = $ref -replace "\d+", ""
    $idx = 0
    foreach ($c in $colStr.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$c - [int][char]'A' + 1)
    }
    return $idx - 1
}

# Target columns for inspection
$targetCols = @(0, 1, 3, 59, 61, 62, 78, 79, 85, 86)  # no, item_code, desc, avg_monthly, total, value, disp, disp_val, stor, stor_val

# Print rows 5-10 (first 6 data rows)
$rowCount = 0
foreach ($row in $rows) {
    $rNum = [int]$row.r
    if ($rNum -lt 5) { continue }
    if ($rNum -gt 10) { break }
    $rowCount++
    Write-Host "`n=== Data Row $rNum ==="
    foreach ($cell in $row.c) {
        $colIdx = Col-Index $cell.r
        if ($colIdx -in $targetCols) {
            $val = Get-CellValue $cell
            Write-Host "  Col $colIdx ($($cell.r -replace '\d+','')): [$val]"
        }
    }
}
Write-Host "`nTotal data rows read: $rowCount"
$zip.Dispose()
