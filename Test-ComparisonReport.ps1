# Test script to generate comparison report using test data
# This simulates the comparison without connecting to SharePoint

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Create folders for data and reports
$dataFolder = Join-Path $scriptDir "Data"
$reportsFolder = Join-Path $scriptDir "Reports"

if (-not (Test-Path $dataFolder)) {
    New-Item -Path $dataFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created Data folder: $dataFolder" -ForegroundColor Gray
}

if (-not (Test-Path $reportsFolder)) {
    New-Item -Path $reportsFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created Reports folder: $reportsFolder" -ForegroundColor Gray
}

# Generate timestamp for this run
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# Look for combined test CSV file
$testDataFolder = Join-Path $scriptDir "testdata"
$combinedCsv = Join-Path $testDataFolder "test-combined.csv"

if (-not (Test-Path $combinedCsv)) {
    Write-Host "Error: test-combined.csv not found in testdata folder." -ForegroundColor Red
    Write-Host "Expected location: $combinedCsv" -ForegroundColor Yellow
    exit 1
}

# Output path for report with timestamp
$OutputPath = Join-Path $reportsFolder "Test-SharePoint-Comparison_$timestamp.html"

# Import the shared HTML report generator
$reportGenerator = Join-Path $scriptDir "New-SharePointComparisonReport.ps1"
if (-not (Test-Path $reportGenerator)) {
    Write-Host "Error: New-SharePointComparisonReport.ps1 not found in the script directory." -ForegroundColor Red
    Write-Host "Expected location: $reportGenerator" -ForegroundColor Yellow
    exit 1
}
. $reportGenerator

$Site1Name = "Last Iteration Site"
$Site2Name = "Current Iteration Site"
$Site1Url = "https://contoso.sharepoint.com/sites/lastiteration"
$Site2Url = "https://contoso.sharepoint.com/sites/currentiteration"

Write-Host "Loading test data..." -ForegroundColor Cyan

# Import combined CSV and split by SourceSite
$allFiles = Import-Csv $combinedCsv

# Get unique site names from the data
$uniqueSites = $allFiles | Select-Object -ExpandProperty SourceSite -Unique

if ($uniqueSites.Count -lt 2) {
    Write-Host "Warning: Combined CSV should contain data from 2 sites, found $($uniqueSites.Count)" -ForegroundColor Yellow
}

# Split data by site (use first two unique sites if names don't match exactly)
$site1Files = $allFiles | Where-Object { $_.SourceSite -eq $uniqueSites[0] }
$site2Files = $allFiles | Where-Object { $_.SourceSite -eq $uniqueSites[1] }

# Update site names and URLs from data if available
if ($uniqueSites.Count -ge 2) {
    $Site1Name = $uniqueSites[0]
    $Site2Name = $uniqueSites[1]
    $Site1Url = ($site1Files | Select-Object -First 1).SourceUrl
    $Site2Url = ($site2Files | Select-Object -First 1).SourceUrl
}

Write-Host "  ${Site1Name}: $($site1Files.Count) files" -ForegroundColor Gray
Write-Host "  ${Site2Name}: $($site2Files.Count) files" -ForegroundColor Gray

Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow

# Generate the HTML report using the shared function
$reportPath = New-SharePointComparisonReport `
    -Site1Files $site1Files `
    -Site2Files $site2Files `
    -Site1Name $Site1Name `
    -Site2Name $Site2Name `
    -Site1Url $Site1Url `
    -Site2Url $Site2Url `
    -OutputPath $OutputPath

Write-Host "`nReport generated successfully!" -ForegroundColor Green
Write-Host "Location: $reportPath" -ForegroundColor Cyan
Write-Host "`nOpening report in default browser..." -ForegroundColor Gray

# Open the report in default browser
Start-Process $reportPath