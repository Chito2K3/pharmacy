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

function Col-Letter($ref) {
    # Extract ONLY the letter part from cell reference like "BJ4" → "BJ"
    return ($ref -replace "\d+$", "")
}

function Col-Index($ref) {
    $colStr = Col-Letter $ref
    $idx = 0
    foreach ($c in $colStr.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$c - [int][char]'A' + 1)
    }
    return $idx - 1
}

# Print row 4 headers in full, showing column letter and index
Write-Host "=== ALL Row 4 Column Headers ==="
foreach ($row in $rows) {
    if ([int]$row.r -ne 4) { continue }
    foreach ($cell in $row.c) {
        $colIdx = Col-Index $cell.r
        $colLetter = Col-Letter $cell.r
        $val = Get-CellValue $cell
        if ($val -ne "") {
            # Truncate long text
            $display = if ($val.Length -gt 80) { $val.Substring(0,80) + "..." } else { $val }
            Write-Host "  Idx=$colIdx ($colLetter): $display"
        }
    }
    break
}
$zip.Dispose()
