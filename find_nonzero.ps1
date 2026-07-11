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

# For each column, track which have ANY non-zero value in data rows (row 5+)
$colSums = @{}
$colNonZero = @{}
$rowCount = 0

foreach ($row in $rows) {
    $rNum = [int]$row.r
    if ($rNum -lt 5) { continue }
    if ($rNum -gt 50) { break }  # check first 46 data rows
    $rowCount++
    foreach ($cell in $row.c) {
        $colIdx = Col-Index $cell.r
        $val = Get-CellValue $cell
        if ($val -match "^[\d\.]+$") {
            $numVal = [double]$val
            if (-not $colSums.ContainsKey($colIdx)) { $colSums[$colIdx] = 0; $colNonZero[$colIdx] = 0 }
            $colSums[$colIdx] += $numVal
            if ($numVal -ne 0) { $colNonZero[$colIdx]++ }
        }
    }
}

Write-Host "Columns with non-zero values in data rows 5-50:"
$colNonZero.GetEnumerator() | Where-Object { $_.Value -gt 0 } | Sort-Object Name | ForEach-Object {
    $col = $_.Name
    $letters = ""
    $n = $col + 1
    while ($n -gt 0) {
        $n--
        $letters = [char]([int][char]'A' + ($n % 26)) + $letters
        $n = [int]($n / 26)
    }
    Write-Host "  Col $col ($letters): $($_.Value) non-zero rows, sum=$($colSums[$col])"
}
$zip.Dispose()
