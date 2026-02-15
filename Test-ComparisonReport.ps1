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

# Look for existing test CSV files or use the ones in testdata folder
$testDataFolder = Join-Path $scriptDir "testdata"
if (Test-Path $testDataFolder) {
    $site1Csv = Join-Path $testDataFolder "test-site1.csv"
    $site2Csv = Join-Path $testDataFolder "test-site2.csv"
} else {
    # Fall back to root folder for backward compatibility
    $site1Csv = Join-Path $scriptDir "test-site1.csv"
    $site2Csv = Join-Path $scriptDir "test-site2.csv"
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

$Site1Name = "Current Iteration Site"
$Site2Name = "Last Iteration Site"
$Site1Url = "https://contoso.sharepoint.com/sites/production"
$Site2Url = "https://contoso.sharepoint.com/sites/archive"

Write-Host "Loading test data..." -ForegroundColor Cyan

$testdataFolder = Join-Path $scriptDir "testdata"
if (-not (Test-Path $testdataFolder)) {
    Write-Host "Error: testdata folder not found in the script directory." -ForegroundColor Red
    Write-Host "Expected location: $testdataFolder" -ForegroundColor Yellow
    exit 1
}
$site1Files = Import-Csv (Join-Path $testdataFolder "test-site1.csv")
$site2Files = Import-Csv (Join-Path $testdataFolder "test-site2.csv")

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