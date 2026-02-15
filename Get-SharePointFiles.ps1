# SharePoint Site File Inventory Script
# This script connects to a SharePoint Online site and exports all file names to a CSV

<#
.SYNOPSIS
    Traverses a SharePoint site and lists all files to a CSV file.

.DESCRIPTION
    This script uses PnP.PowerShell to connect to SharePoint Online,
    recursively traverse all document libraries and folders, and export
    file information to a CSV file.

.PARAMETER SiteUrl
    The URL of the SharePoint site (e.g., https://yourtenant.sharepoint.com/sites/yoursite)

.PARAMETER OutputPath
    The path where the CSV file will be saved (default: SharePointFiles.csv)

.EXAMPLE
    .\Get-SharePointFiles.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite" -OutputPath "files.csv"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "SharePointFiles.csv",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipConnection
)

# Initialize SharePoint connection (unless skipped by caller)
if (-not $SkipConnection) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $initScript = Join-Path $scriptDir "Initialize-SharePointConnection.ps1"
    
    if (-not (Test-Path $initScript)) {
        Write-Host "Error: Initialize-SharePointConnection.ps1 not found." -ForegroundColor Red
        Write-Host "Expected location: $initScript" -ForegroundColor Yellow
        exit 1
    }
    
    # Dot-source the initialization script and call the function
    . $initScript
    
    try {
        Initialize-SharePointConnection -SiteUrl $SiteUrl
    }
    catch {
        Write-Host "Failed to initialize SharePoint connection." -ForegroundColor Red
        exit 1
    }
}

# Array to store file information
$fileList = @()

# Function to recursively get files from a folder
function Get-FilesRecursively {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderUrl,
        
        [Parameter(Mandatory=$false)]
        [string]$LibraryName = ""
    )
    
    try {
        # Get all items in the current folder
        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderUrl -ItemType File
        
        foreach ($item in $items) {
            $fileInfo = [PSCustomObject]@{
                FileName = $item.Name
                FilePath = $item.ServerRelativeUrl
                FileSize = $item.Length
                Created = $item.TimeCreated
                Modified = $item.TimeLastModified
                Library = $LibraryName
                FileExtension = $item.Name.Substring($item.Name.LastIndexOf('.') + 1)
            }
            $script:fileList += $fileInfo
            Write-Host "Found: $($item.ServerRelativeUrl)" -ForegroundColor Gray
        }
        
        # Get all subfolders and recurse
        $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderUrl -ItemType Folder
        
        foreach ($folder in $folders) {
            Get-FilesRecursively -FolderUrl $folder.ServerRelativeUrl -LibraryName $LibraryName
        }
    }
    catch {
        Write-Host "Error accessing folder $FolderUrl : $_" -ForegroundColor Red
    }
}

try {
    # Connection is already established by Initialize-SharePointConnection
    # Unless SkipConnection was specified (for when called from Compare script)
    if ($SkipConnection) {
        Write-Host "Using existing SharePoint connection..." -ForegroundColor Cyan
    }
    
    Write-Host "Retrieving document libraries..." -ForegroundColor Cyan
    
    # Get all document libraries
    $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
    
    Write-Host "Found $($lists.Count) document libraries. Starting file enumeration..." -ForegroundColor Cyan
    
    # Process each document library
    foreach ($list in $lists) {
        Write-Host "`nProcessing library: $($list.Title)" -ForegroundColor Yellow
        
        $rootFolder = $list.RootFolder.ServerRelativeUrl
        Get-FilesRecursively -FolderUrl $rootFolder -LibraryName $list.Title
    }
    
    # Export to CSV
    if ($fileList.Count -gt 0) {
        $fileList | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nSuccess! Found $($fileList.Count) files." -ForegroundColor Green
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
    }
    else {
        Write-Host "`nNo files found in the SharePoint site." -ForegroundColor Yellow
    }
    
    # Disconnect (only if running standalone, not when called from Compare script)
    if (-not $SkipConnection) {
        Disconnect-PnPOnline
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    # Ensure we disconnect even if there's an error (only for standalone runs)
    if (-not $SkipConnection) {
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    }
}
