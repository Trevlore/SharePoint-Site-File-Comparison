# SharePoint Connection Initialization Script
# Ensures PnP PowerShell is installed and handles authentication with 2FA support

<#
.SYNOPSIS
    Initializes and authenticates a connection to SharePoint Online.

.DESCRIPTION
    This script ensures that the PnP.PowerShell module is installed and imported,
    then establishes an authenticated connection to a SharePoint site with support
    for multi-factor authentication (MFA/2FA).

.PARAMETER SiteUrl
    The URL of the SharePoint site to connect to

.PARAMETER Force
    Force reinstallation of PnP.PowerShell module

.EXAMPLE
    .\Initialize-SharePointConnection.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite"

.EXAMPLE
    # Use as a function by dot-sourcing
    . .\Initialize-SharePointConnection.ps1
    Initialize-SharePointConnection -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

function Initialize-SharePointConnection {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "SharePoint Connection Initialization" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    
    # Step 1: Check if PnP.PowerShell module is installed
    Write-Host "`n[1/3] Checking PnP.PowerShell module..." -ForegroundColor Yellow
    
    $pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell
    
    if (-not $pnpModule -or $Force) {
        if ($Force) {
            Write-Host "  Force flag set - reinstalling PnP.PowerShell..." -ForegroundColor Yellow
        } else {
            Write-Host "  PnP.PowerShell module not found." -ForegroundColor Yellow
        }
        
        try {
            Write-Host "  Installing PnP.PowerShell module..." -ForegroundColor Cyan
            Write-Host "  (This may take a few minutes on first install)" -ForegroundColor Gray
            
            # Uninstall old version if Force is specified
            if ($Force -and $pnpModule) {
                Write-Host "  Removing existing version..." -ForegroundColor Gray
                Uninstall-Module -Name PnP.PowerShell -AllVersions -Force -ErrorAction SilentlyContinue
            }
            
            # Install the module
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
            
            Write-Host "  ✓ PnP.PowerShell module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Failed to install PnP.PowerShell module" -ForegroundColor Red
            Write-Host "  Error: $_" -ForegroundColor Red
            throw "PnP.PowerShell installation failed"
        }
    } else {
        Write-Host "  ✓ PnP.PowerShell module found (Version: $($pnpModule[0].Version))" -ForegroundColor Green
    }
    
    # Step 2: Import the module
    Write-Host "`n[2/3] Importing PnP.PowerShell module..." -ForegroundColor Yellow
    
    try {
        # Remove module first if it's already loaded to ensure fresh import
        if (Get-Module -Name PnP.PowerShell) {
            Remove-Module -Name PnP.PowerShell -Force -ErrorAction SilentlyContinue
        }
        
        Import-Module PnP.PowerShell -DisableNameChecking -WarningAction SilentlyContinue
        Write-Host "  ✓ Module imported successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ Failed to import PnP.PowerShell module" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        throw "PnP.PowerShell import failed"
    }
    
    # Step 3: Connect to SharePoint with Interactive authentication (2FA supported)
    Write-Host "`n[3/3] Connecting to SharePoint site..." -ForegroundColor Yellow
    Write-Host "  Site URL: $SiteUrl" -ForegroundColor Gray
    Write-Host "  Authentication: Interactive (supports MFA/2FA)" -ForegroundColor Gray
    
    try {
        # Check if already connected to this site
        $existingConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        
        if ($existingConnection -and $existingConnection.Url -eq $SiteUrl) {
            Write-Host "  ℹ Already connected to $SiteUrl" -ForegroundColor Cyan
            
            # Verify connection is active
            try {
                $testWeb = Get-PnPWeb -ErrorAction Stop
                Write-Host "  ✓ Connection verified - you are connected as: $($existingConnection.PSCredential.UserName)" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "  ⚠ Existing connection not active, reconnecting..." -ForegroundColor Yellow
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
            }
        }
        elseif ($existingConnection) {
            Write-Host "  ℹ Disconnecting from current site: $($existingConnection.Url)" -ForegroundColor Gray
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        
        Write-Host "`n  Please authenticate when prompted..." -ForegroundColor Cyan
        Write-Host "  → A browser window will open for authentication" -ForegroundColor Gray
        Write-Host "  → Complete MFA/2FA if required" -ForegroundColor Gray
        Write-Host "  → You may need to consent to app permissions on first use" -ForegroundColor Gray
        Write-Host "" # Blank line for readability
        
        # Connect using Interactive authentication (supports MFA/2FA)
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        
        # Verify the connection
        $connection = Get-PnPConnection
        $web = Get-PnPWeb
        
        Write-Host "`n  ✓ Connected successfully!" -ForegroundColor Green
        Write-Host "  Site: $($web.Title)" -ForegroundColor Gray
        Write-Host "  URL: $($web.Url)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Host "`n  ✗ Failed to connect to SharePoint site" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        
        # Provide helpful troubleshooting tips
        Write-Host "`n  Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  • Ensure you have access to the SharePoint site" -ForegroundColor Gray
        Write-Host "  • Check that the site URL is correct" -ForegroundColor Gray
        Write-Host "  • Verify your internet connection" -ForegroundColor Gray
        Write-Host "  • If using MFA, ensure authentication completes in browser" -ForegroundColor Gray
        Write-Host "  • You may need admin consent for PnP app permissions" -ForegroundColor Gray
        
        throw "SharePoint connection failed"
    }
}

# If SiteUrl is provided as a parameter, execute the function
if (-not [string]::IsNullOrEmpty($SiteUrl)) {
    try {
        $result = Initialize-SharePointConnection -SiteUrl $SiteUrl -Force:$Force
        exit 0
    }
    catch {
        exit 1
    }
}
else {
    # Script was dot-sourced, function is now available
    Write-Host "Initialize-SharePointConnection function loaded." -ForegroundColor Green
    Write-Host "Usage: Initialize-SharePointConnection -SiteUrl 'https://tenant.sharepoint.com/sites/site'" -ForegroundColor Gray
}
