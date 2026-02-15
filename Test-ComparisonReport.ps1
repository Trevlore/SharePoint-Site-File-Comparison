# Test script to generate comparison report using test data
# This simulates the comparison without connecting to SharePoint

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$site1Csv = Join-Path $scriptDir "test-site1.csv"
$site2Csv = Join-Path $scriptDir "test-site2.csv"
$OutputPath = Join-Path $scriptDir "Test-SharePoint-Comparison-Report.html"

# Import the shared HTML report generator
$reportGenerator = Join-Path $scriptDir "New-SharePointComparisonReport.ps1"
if (-not (Test-Path $reportGenerator)) {
    Write-Host "Error: New-SharePointComparisonReport.ps1 not found in the script directory." -ForegroundColor Red
    Write-Host "Expected location: $reportGenerator" -ForegroundColor Yellow
    exit 1
}
. $reportGenerator

$Site1Name = "Production Site"
$Site2Name = "Archive Site"
$Site1Url = "https://contoso.sharepoint.com/sites/production"
$Site2Url = "https://contoso.sharepoint.com/sites/archive"

Write-Host "Loading test data..." -ForegroundColor Cyan

$site1Files = Import-Csv $site1Csv
$site2Files = Import-Csv $site2Csv

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

Write-Host "`nReport generated successfully!" -ForegroundColor Green
Write-Host "Location: $OutputPath" -ForegroundColor Cyan
Write-Host "`nOpening report in default browser..." -ForegroundColor Gray

# Open the report in default browser
Start-Process $OutputPath
