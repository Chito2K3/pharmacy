# Convert / update locator dataset dynamically relative to project directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = (Get-Item $scriptDir).Parent.FullName

$locJsonPath = Join-Path $projectDir "locator.json"
$locJsPath   = Join-Path $projectDir "locator_data.js"

if (-not (Test-Path $locJsonPath)) {
    Write-Error "locator.json not found at $locJsonPath"
    exit 1
}

Write-Host "Reading locator dataset from $locJsonPath..."
$items = Get-Content -LiteralPath $locJsonPath -Raw | ConvertFrom-Json

# Apply location rules / updates
foreach ($item in $items) {
    if ($item.item_code -match "0213" -and $item.item_code -match "DMR") {
        $item.location = "L6"
    }
}

$json = $items | ConvertTo-Json -Depth 4
$json | Set-Content -Path $locJsonPath -Encoding UTF8

$jsContent = "window.LOCATOR_DATA = " + $json + ";"
$jsContent | Set-Content -Path $locJsPath -Encoding UTF8

Write-Host "Successfully updated locator.json and locator_data.js."
