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

# Load shared strings
$ssXml = Get-EntryXML $zip "xl/sharedStrings.xml"
$sharedStrings = @()
if ($ssXml) {
    $sharedStrings = $ssXml.sst.si | ForEach-Object {
        if ($_.t) { $_.t }
        elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' }
        else { '' }
    }
}

# sheet3 = Inventory Utilization 2025
$sheetXml = Get-EntryXML $zip "xl/worksheets/sheet3.xml"
$rows = $sheetXml.worksheet.sheetData.row

function Get-CellValue($cell) {
    $v = $cell.v
    if ($null -eq $v) { return "" }
    if ($cell.t -eq "s") { return $sharedStrings[[int]$v] }
    return $v
}

function Col-Index($ref) {
    # Convert e.g. "BJ5" -> column index (0-based)
    $colStr = $ref -replace "\d+", ""
    $idx = 0
    foreach ($c in $colStr.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$c - [int][char]'A' + 1)
    }
    return $idx - 1
}

# Print first 4 rows to see header structure
$rowCount = 0
foreach ($row in $rows) {
    $rowCount++
    if ($rowCount -gt 4) { break }
    Write-Host "=== Row $($row.r) ==="
    foreach ($cell in $row.c) {
        $colIdx = Col-Index $cell.r
        $val = Get-CellValue $cell
        if ($val -ne "" -and $colIdx -ge 55) {
            Write-Host "  Col $colIdx ($($cell.r -replace '\d+','')): $val"
        }
    }
}
$zip.Dispose()
