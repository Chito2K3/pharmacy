[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$xlsxPath = "MASTERLIST_2026_V30.xlsx"
$jsonPath = "items.json"

if (-not (Test-Path $xlsxPath)) {
    Write-Error "Excel file not found: $xlsxPath"
    exit 1
}
if (-not (Test-Path $jsonPath)) {
    Write-Error "JSON file not found: $jsonPath"
    exit 1
}

Write-Host "Opening $xlsxPath..."
$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)

# 1. Fast stream read shared strings
Write-Host "Reading shared strings..."
$entry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
$stream = $entry.Open()
$reader = [System.Xml.XmlReader]::Create($stream)
$strings = New-Object System.Collections.Generic.List[string]
while ($reader.Read()) {
    if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $reader.Name -eq "t") {
        $strings.Add($reader.ReadElementContentAsString())
    }
}
$reader.Close(); $stream.Close()
Write-Host "Read $($strings.Count) shared strings."

# 2. Read items.json
Write-Host "Reading $jsonPath..."
$itemsJsonText = [System.IO.File]::ReadAllText($jsonPath)
$items = $itemsJsonText | ConvertFrom-Json
$itemMap = @{}
foreach ($it in $items) {
    if ($it.item_code) { $itemMap[$it.item_code.Trim()] = $it }
    if ($it.dci_code) { $itemMap[$it.dci_code.Trim()] = $it }
}
Write-Host "Indexed $($itemMap.Count) keys from $($items.Count) items."

# 3. Read sheet3.xml (Inventory Utilization 2025)
Write-Host "Parsing sheet3.xml for monthly consumption (Col AI-AT)..."
$entry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet3.xml" }
$sStream = $entry.Open()
$sReader = [System.Xml.XmlReader]::Create($sStream)

function Col-Index($ref) {
    $colStr = $ref -replace "\d+$", ""
    $idx = 0
    foreach ($c in $colStr.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$c - [int][char]'A' + 1)
    }
    return $idx - 1
}

$currentRowCode = ""
$currentMonthly = @(0,0,0,0,0,0,0,0,0,0,0,0)
$hasMonthly = $false
$enrichedCount = 0

while ($sReader.Read()) {
    if ($sReader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $sReader.Name -eq "row") {
        if ($hasMonthly -and $currentRowCode) {
            if ($itemMap.ContainsKey($currentRowCode)) {
                $target = $itemMap[$currentRowCode]
                $target | Add-Member -NotePropertyName "monthly_consumption" -NotePropertyValue $currentMonthly -Force
                $enrichedCount++
            }
        }
        $currentRowCode = ""
        $currentMonthly = @(0,0,0,0,0,0,0,0,0,0,0,0)
        $hasMonthly = $false
    }
    elseif ($sReader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $sReader.Name -eq "c") {
        $r = $sReader.GetAttribute("r")
        $t = $sReader.GetAttribute("t")
        $colIdx = Col-Index $r
        
        $cellSub = $sReader.ReadSubtree()
        while ($cellSub.Read()) {
            if ($cellSub.NodeType -eq [System.Xml.XmlNodeType]::Element -and $cellSub.Name -eq "v") {
                $vStr = $cellSub.ReadElementContentAsString()
                if ($colIdx -eq 1) { # Col B = Item Code
                    if ($t -eq "s") {
                        $sIdx = [int]$vStr
                        if ($sIdx -lt $strings.Count) { $currentRowCode = $strings[$sIdx].Trim() }
                    } else {
                        $currentRowCode = $vStr.Trim()
                    }
                }
                elseif ($colIdx -ge 34 -and $colIdx -le 45) { # Col AI to AT = Jan-Dec 2025
                    $mIdx = $colIdx - 34
                    $numVal = 0.0
                    if ([double]::TryParse($vStr, [ref]$numVal)) {
                        if ($numVal -gt 0) {
                            $currentMonthly[$mIdx] = [math]::Round($numVal, 2)
                            $hasMonthly = $true
                        }
                    }
                }
            }
        }
        $cellSub.Close()
    }
}
$sReader.Close(); $sStream.Close()
$zip.Dispose()

Write-Host "Enriched $enrichedCount items with monthly consumption data."

# 4. Save back to items.json
Write-Host "Writing updated items to $jsonPath..."
$newJson = $items | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($jsonPath, $newJson)

$sw.Stop()
Write-Host "Sync completed successfully in $($sw.ElapsedMilliseconds) ms!" -ForegroundColor Green
