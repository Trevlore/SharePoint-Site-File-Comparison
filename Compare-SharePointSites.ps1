# SharePoint Sites Comparison Script
# This script compares files between two SharePoint sites and generates an HTML report

<#
.SYNOPSIS
    Compares files between two SharePoint sites and produces an HTML report.

.DESCRIPTION
    This script uses Get-SharePointFiles.ps1 to retrieve file inventories from two
    SharePoint sites, compares them, and generates a formatted HTML report showing
    files unique to each site and files common to both.

.PARAMETER Site1Url
    The URL of the first SharePoint site

.PARAMETER Site2Url
    The URL of the second SharePoint site

.PARAMETER Site1Name
    Display name for the first site (default: "Site 1")

.PARAMETER Site2Name
    Display name for the second site (default: "Site 2")

.PARAMETER OutputPath
    Path for the HTML report (default: SharePoint-Comparison-Report.html)

.EXAMPLE
    .\Compare-SharePointSites.ps1 -Site1Url "https://contoso.sharepoint.com/sites/TeamSite1" -Site2Url "https://contoso.sharepoint.com/sites/TeamSite2"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Site1Url,
    
    [Parameter(Mandatory=$true)]
    [string]$Site2Url,
    
    [Parameter(Mandatory=$false)]
    [string]$Site1Name = "Site 1",
    
    [Parameter(Mandatory=$false)]
    [string]$Site2Name = "Site 2",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ""
)

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$getFilesScript = Join-Path $scriptDir "Get-SharePointFiles.ps1"
$initScript = Join-Path $scriptDir "Initialize-SharePointConnection.ps1"

# Check if required scripts exist
if (-not (Test-Path $getFilesScript)) {
    Write-Host "Error: Get-SharePointFiles.ps1 not found in the script directory." -ForegroundColor Red
    Write-Host "Expected location: $getFilesScript" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $initScript)) {
    Write-Host "Error: Initialize-SharePointConnection.ps1 not found in the script directory." -ForegroundColor Red
    Write-Host "Expected location: $initScript" -ForegroundColor Yellow
    exit 1
}

# Load the initialization script function
. $initScript

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

# Temporary CSV files for individual site data
$site1Csv = Join-Path $env:TEMP "temp_site1_$timestamp.csv"
$site2Csv = Join-Path $env:TEMP "temp_site2_$timestamp.csv"

# Combined CSV file with timestamp
$combinedCsv = Join-Path $dataFolder "Comparison_$timestamp.csv"

try {
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "SharePoint Sites Comparison Tool" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    
    # Step 1: Connect to Site 1 and get files
    Write-Host "`n[1/5] Connecting to $Site1Name..." -ForegroundColor Yellow
    Write-Host "Site URL: $Site1Url" -ForegroundColor Gray
    
    try {
        Initialize-SharePointConnection -SiteUrl $Site1Url
    }
    catch {
        throw "Failed to connect to ${Site1Name}"
    }
    
    Write-Host "`n[2/5] Retrieving files from $Site1Name..." -ForegroundColor Yellow
    & $getFilesScript -SiteUrl $Site1Url -OutputPath $site1Csv -SkipConnection
    
    if (-not (Test-Path $site1Csv)) {
        throw "Failed to retrieve files from $Site1Name"
    }
    
    # Disconnect from Site 1
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    
    # Step 2: Connect to Site 2 and get files
    Write-Host "`n[3/5] Connecting to $Site2Name..." -ForegroundColor Yellow
    Write-Host "Site URL: $Site2Url" -ForegroundColor Gray
    
    try {
        Initialize-SharePointConnection -SiteUrl $Site2Url
    }
    catch {
        throw "Failed to connect to ${Site2Name}"
    }
    
    Write-Host "`n[4/5] Retrieving files from $Site2Name..." -ForegroundColor Yellow
    & $getFilesScript -SiteUrl $Site2Url -OutputPath $site2Csv -SkipConnection
    
    if (-not (Test-Path $site2Csv)) {
        throw "Failed to retrieve files from $Site2Name"
    }
    
    # Disconnect from Site 2
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    
    # Step 3: Combine data into single CSV
    Write-Host "`n[5/5] Combining and saving data..." -ForegroundColor Yellow
    
    $site1Files = Import-Csv $site1Csv
    $site2Files = Import-Csv $site2Csv
    
    Write-Host "  ${Site1Name}: $($site1Files.Count) files" -ForegroundColor Gray
    Write-Host "  ${Site2Name}: $($site2Files.Count) files" -ForegroundColor Gray
    
    # Add SourceSite column to each dataset
    $site1FilesWithSource = $site1Files | Select-Object *, @{Name='SourceSite';Expression={$Site1Name}}, @{Name='SourceUrl';Expression={$Site1Url}}
    $site2FilesWithSource = $site2Files | Select-Object *, @{Name='SourceSite';Expression={$Site2Name}}, @{Name='SourceUrl';Expression={$Site2Url}}
    
    # Combine into single array
    $allFiles = @()
    $allFiles += $site1FilesWithSource
    $allFiles += $site2FilesWithSource
    
    # Save combined data to CSV
    $allFiles | Export-Csv -Path $combinedCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  Combined data saved: $combinedCsv" -ForegroundColor Gray
    
    # Step 4: Analyze comparison data
    Write-Host "`nAnalyzing comparison..." -ForegroundColor Yellow
    
    # Create hashtables for quick lookup (using relative path as key)
    $site1Hash = @{}
    foreach ($file in $site1Files) {
        # Use FilePath as unique identifier
        $site1Hash[$file.FilePath] = $file
    }
    
    $site2Hash = @{}
    foreach ($file in $site2Files) {
        $site2Hash[$file.FilePath] = $file
    }
    
    # Find files only in Site 1
    $onlyInSite1 = @()
    foreach ($file in $site1Files) {
        if (-not $site2Hash.ContainsKey($file.FilePath)) {
            $onlyInSite1 += $file
        }
    }
    
    # Find files only in Site 2
    $onlyInSite2 = @()
    foreach ($file in $site2Files) {
        if (-not $site1Hash.ContainsKey($file.FilePath)) {
            $onlyInSite2 += $file
        }
    }
    
    # Find files in both sites
    $inBothSites = @()
    foreach ($file in $site1Files) {
        if ($site2Hash.ContainsKey($file.FilePath)) {
            $site2File = $site2Hash[$file.FilePath]
            $comparisonObj = [PSCustomObject]@{
                FileName = $file.FileName
                FilePath = $file.FilePath
                Library = $file.Library
                Site1Size = $file.FileSize
                Site2Size = $site2File.FileSize
                Site1Modified = $file.Modified
                Site2Modified = $site2File.Modified
                SizeMatch = ($file.FileSize -eq $site2File.FileSize)
                ModifiedMatch = ($file.Modified -eq $site2File.Modified)
            }
            $inBothSites += $comparisonObj
        }
    }
    
    Write-Host "  Only in ${Site1Name}: $($onlyInSite1.Count) files" -ForegroundColor Cyan
    Write-Host "  Only in ${Site2Name}: $($onlyInSite2.Count) files" -ForegroundColor Cyan
    Write-Host "  In both sites: $($inBothSites.Count) files" -ForegroundColor Cyan
    
    # Step 5: Generate HTML report
    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    
    # Set output path with timestamp if not provided
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Join-Path $reportsFolder "SharePoint-Comparison_$timestamp.html"
    } elseif (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
        # If relative path provided, put it in reports folder
        $OutputPath = Join-Path $reportsFolder $OutputPath
    }
    
    $reportDate = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint Sites Comparison Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        
        .header .subtitle {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .summary {
            background: #f8f9fa;
            padding: 30px;
            border-bottom: 3px solid #667eea;
        }
        
        .summary h2 {
            color: #667eea;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 4px solid #667eea;
        }
        
        .stat-card h3 {
            color: #666;
            font-size: 0.9em;
            text-transform: uppercase;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .stat-card .value {
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }
        
        .stat-card .label {
            color: #999;
            font-size: 0.9em;
            margin-top: 5px;
        }
        
        .site-info {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-top: 20px;
        }
        
        .site-box {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .site-box h3 {
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .site-box .url {
            color: #666;
            font-size: 0.9em;
            word-break: break-all;
        }
        
        .section {
            padding: 30px;
        }
        
        .section h2 {
            color: #667eea;
            margin-bottom: 20px;
            font-size: 1.8em;
            padding-bottom: 10px;
            border-bottom: 2px solid #e0e0e0;
        }
        
        .section-description {
            color: #666;
            margin-bottom: 20px;
            font-style: italic;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        
        thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        tbody tr:hover {
            background: #f8f9fa;
            transition: background 0.2s ease;
        }
        
        tbody tr:last-child td {
            border-bottom: none;
        }
        
        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: 600;
        }
        
        .badge-success {
            background: #d4edda;
            color: #155724;
        }
        
        .badge-warning {
            background: #fff3cd;
            color: #856404;
        }
        
        .badge-danger {
            background: #f8d7da;
            color: #721c24;
        }
        
        .no-data {
            text-align: center;
            padding: 40px;
            color: #999;
            font-style: italic;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #666;
            font-size: 0.9em;
            border-top: 2px solid #e0e0e0;
        }
        
        .highlight-different {
            background: #fff3cd;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 SharePoint Sites Comparison Report</h1>
            <div class="subtitle">Generated on $reportDate</div>
        </div>
        
        <div class="summary">
            <h2>Summary</h2>
            
            <div class="site-info">
                <div class="site-box">
                    <h3>$Site1Name</h3>
                    <div class="url">$Site1Url</div>
                    <div class="value" style="margin-top: 10px; color: #667eea;">$($site1Files.Count) files</div>
                </div>
                <div class="site-box">
                    <h3>$Site2Name</h3>
                    <div class="url">$Site2Url</div>
                    <div class="value" style="margin-top: 10px; color: #667eea;">$($site2Files.Count) files</div>
                </div>
            </div>
            
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>Only in $Site1Name</h3>
                    <div class="value">$($onlyInSite1.Count)</div>
                    <div class="label">unique files</div>
                </div>
                <div class="stat-card">
                    <h3>Only in $Site2Name</h3>
                    <div class="value">$($onlyInSite2.Count)</div>
                    <div class="label">unique files</div>
                </div>
                <div class="stat-card">
                    <h3>Common Files</h3>
                    <div class="value">$($inBothSites.Count)</div>
                    <div class="label">files in both sites</div>
                </div>
                <div class="stat-card">
                    <h3>Total Unique</h3>
                    <div class="value">$($site1Hash.Count + $site2Hash.Count - $inBothSites.Count)</div>
                    <div class="label">distinct file paths</div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2>Files Only in $Site1Name</h2>
            <div class="section-description">These files exist in $Site1Name but not in $Site2Name (based on file path)</div>
"@
    
    if ($onlyInSite1.Count -gt 0) {
        $html += @"
            <table>
                <thead>
                    <tr>
                        <th>File Name</th>
                        <th>File Path</th>
                        <th>Library</th>
                        <th>Size (bytes)</th>
                        <th>Modified</th>
                    </tr>
                </thead>
                <tbody>
"@
        foreach ($file in $onlyInSite1 | Sort-Object FilePath) {
            $html += @"
                    <tr>
                        <td>$($file.FileName)</td>
                        <td>$($file.FilePath)</td>
                        <td>$($file.Library)</td>
                        <td>$($file.FileSize)</td>
                        <td>$($file.Modified)</td>
                    </tr>
"@
        }
        $html += @"
                </tbody>
            </table>
"@
    } else {
        $html += '<div class="no-data">No unique files found</div>'
    }
    
    $html += @"
        </div>
        
        <div class="section">
            <h2>Files Only in $Site2Name</h2>
            <div class="section-description">These files exist in $Site2Name but not in $Site1Name (based on file path)</div>
"@
    
    if ($onlyInSite2.Count -gt 0) {
        $html += @"
            <table>
                <thead>
                    <tr>
                        <th>File Name</th>
                        <th>File Path</th>
                        <th>Library</th>
                        <th>Size (bytes)</th>
                        <th>Modified</th>
                    </tr>
                </thead>
                <tbody>
"@
        foreach ($file in $onlyInSite2 | Sort-Object FilePath) {
            $html += @"
                    <tr>
                        <td>$($file.FileName)</td>
                        <td>$($file.FilePath)</td>
                        <td>$($file.Library)</td>
                        <td>$($file.FileSize)</td>
                        <td>$($file.Modified)</td>
                    </tr>
"@
        }
        $html += @"
                </tbody>
            </table>
"@
    } else {
        $html += '<div class="no-data">No unique files found</div>'
    }
    
    $html += @"
        </div>
        
        <div class="section">
            <h2>Common Files (In Both Sites)</h2>
            <div class="section-description">These files exist in both sites. Rows highlighted indicate differences in size or modification date.</div>
"@
    
    if ($inBothSites.Count -gt 0) {
        $html += @"
            <table>
                <thead>
                    <tr>
                        <th>File Name</th>
                        <th>File Path</th>
                        <th>Library</th>
                        <th>$Site1Name Size</th>
                        <th>$Site2Name Size</th>
                        <th>$Site1Name Modified</th>
                        <th>$Site2Name Modified</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
"@
        foreach ($file in $inBothSites | Sort-Object FilePath) {
            $rowClass = ""
            $status = ""
            if (-not $file.SizeMatch -or -not $file.ModifiedMatch) {
                $rowClass = ' class="highlight-different"'
                if (-not $file.SizeMatch -and -not $file.ModifiedMatch) {
                    $status = '<span class="badge badge-warning">Size & Date Differ</span>'
                } elseif (-not $file.SizeMatch) {
                    $status = '<span class="badge badge-warning">Size Differs</span>'
                } else {
                    $status = '<span class="badge badge-warning">Date Differs</span>'
                }
            } else {
                $status = '<span class="badge badge-success">Match</span>'
            }
            
            $html += @"
                    <tr$rowClass>
                        <td>$($file.FileName)</td>
                        <td>$($file.FilePath)</td>
                        <td>$($file.Library)</td>
                        <td>$($file.Site1Size)</td>
                        <td>$($file.Site2Size)</td>
                        <td>$($file.Site1Modified)</td>
                        <td>$($file.Site2Modified)</td>
                        <td>$status</td>
                    </tr>
"@
        }
        $html += @"
                </tbody>
            </table>
"@
    } else {
        $html += '<div class="no-data">No common files found</div>'
    }
    
    $html += @"
        </div>
        
        <div class="footer">
            <p>Report generated by Compare-SharePointSites.ps1</p>
            <p>Comparison based on file paths between two SharePoint sites</p>
        </div>
    </div>
</body>
</html>
"@
    
    # Save the HTML report
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Host "`n==========================================" -ForegroundColor Green
    Write-Host "Report generated successfully!" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "Report: $OutputPath" -ForegroundColor Cyan
    Write-Host "Data: $combinedCsv" -ForegroundColor Cyan
    Write-Host "  - ${Site1Name}: $($site1Files.Count) files" -ForegroundColor Gray
    Write-Host "  - ${Site2Name}: $($site2Files.Count) files" -ForegroundColor Gray
    Write-Host "`nOpening report in default browser..." -ForegroundColor Gray
    
    # Open the report in default browser
    Start-Process $OutputPath
    
} catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
} finally {
    # Clean up temporary CSV files
    if (Test-Path $site1Csv) {
        Remove-Item $site1Csv -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $site2Csv) {
        Remove-Item $site2Csv -Force -ErrorAction SilentlyContinue
    }
    
    # Ensure disconnection from SharePoint
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    } catch {}
    
    # Combined CSV remains in Data folder for historical tracking
}
