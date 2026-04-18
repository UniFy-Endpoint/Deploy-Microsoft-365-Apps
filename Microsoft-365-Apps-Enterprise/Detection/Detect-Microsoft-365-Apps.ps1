<#
.SYNOPSIS
    Detection script for Microsoft 365 Apps deployment via Intune Win32 App

.DESCRIPTION
    This script detects the installation status of Microsoft 365 Apps (O365ProPlusRetail or O365BusinessRetail)
    for use as a detection script in Microsoft Intune Win32 app deployments.
    
    The script checks for:
    - Office Click-to-Run installation registry keys
    - Specific product IDs (O365ProPlusRetail, O365BusinessRetail)
    - Office application executables
    - Version information and installation status
    
    Exit Codes:
    - 0: Application is installed and detected successfully
    - 1: Application is not installed or detection failed

.NOTES
    Version: 1.0.0
    Author: UniFy-Endpoint
    
    
    Usage in Intune:
    - Detection method: Use PowerShell script
    - Run script as 32-bit PowerShell: No
    - Enforce script signature check: No (unless you sign the script)
#>

[CmdletBinding()]
param()

# Initialize variables
$DetectionResult = $false
$LogOutput = @()

# Function to write detection log (for troubleshooting)
function Write-DetectionLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $LogEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    $script:LogOutput += $LogEntry
    
    # Write to console for debugging (will not affect Intune detection)
    switch ($Level) {
        'Info' { Write-Verbose $Message -Verbose }
        'Warning' { Write-Warning $Message }
        'Error' { Write-Error $Message }
    }
}

# Function to check Office Click-to-Run registry
function Test-OfficeClickToRunRegistry {
    try {
        Write-DetectionLog "Checking Office Click-to-Run registry keys..."
        
        # Primary registry paths for Office Click-to-Run
        $RegistryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
        )
        
        foreach ($RegPath in $RegistryPaths) {
            if (Test-Path $RegPath) {
                Write-DetectionLog "Found Click-to-Run registry path: $RegPath"
                
                # Check for ProductReleaseIds which contains the product information
                $ProductReleaseIds = Get-ItemProperty -Path $RegPath -Name "ProductReleaseIds" -ErrorAction SilentlyContinue
                
                if ($ProductReleaseIds) {
                    $ProductIds = $ProductReleaseIds.ProductReleaseIds
                    Write-DetectionLog "Found ProductReleaseIds: $ProductIds"
                    
                    # Check for target product IDs
                    if ($ProductIds -match "O365ProPlusRetail" -or $ProductIds -match "O365BusinessRetail") {
                        Write-DetectionLog "Target Office 365 product detected in registry"
                        return $true
                    }
                }
                
                # Alternative check - look for VersionToReport
                $VersionInfo = Get-ItemProperty -Path $RegPath -Name "VersionToReport" -ErrorAction SilentlyContinue
                if ($VersionInfo) {
                    Write-DetectionLog "Office version detected: $($VersionInfo.VersionToReport)"
                }
            }
        }
        
        return $false
    }
    catch {
        Write-DetectionLog "Error checking Click-to-Run registry: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to check installed Office products via registry
function Test-OfficeInstalledProducts {
    try {
        Write-DetectionLog "Checking installed Office products..."
        
        # Check Office installation registry
        $OfficeRegPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs"
        )
        
        foreach ($RegPath in $OfficeRegPaths) {
            if (Test-Path $RegPath) {
                $Products = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue
                foreach ($Product in $Products) {
                    $ProductName = $Product.PSChildName
                    Write-DetectionLog "Found installed product: $ProductName"
                    
                    if ($ProductName -eq "O365ProPlusRetail" -or $ProductName -eq "O365BusinessRetail") {
                        Write-DetectionLog "Target Office 365 product found: $ProductName"
                        return $true
                    }
                }
            }
        }
        
        # Alternative method - check Uninstall registry
        $UninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        
        foreach ($UninstallPath in $UninstallPaths) {
            if (Test-Path $UninstallPath) {
                $InstalledPrograms = Get-ChildItem -Path $UninstallPath -ErrorAction SilentlyContinue
                foreach ($Program in $InstalledPrograms) {
                    $ProgramInfo = Get-ItemProperty -Path $Program.PSPath -ErrorAction SilentlyContinue
                    if ($ProgramInfo.DisplayName -like "*Microsoft 365*" -or 
                        $ProgramInfo.DisplayName -like "*Office 365*" -or
                        $ProgramInfo.DisplayName -like "*Microsoft Office*") {
                        Write-DetectionLog "Found Office installation: $($ProgramInfo.DisplayName)"
                        return $true
                    }
                }
            }
        }
        
        return $false
    }
    catch {
        Write-DetectionLog "Error checking installed products: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to check Office application executables
function Test-OfficeExecutables {
    try {
        Write-DetectionLog "Checking Office application executables..."
        
        # Common Office installation paths
        $OfficePaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
            "${env:ProgramFiles}\Microsoft Office\Office16",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
        )
        
        # Core Office applications to check
        $OfficeApps = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE")
        
        foreach ($OfficePath in $OfficePaths) {
            if (Test-Path $OfficePath) {
                Write-DetectionLog "Found Office installation path: $OfficePath"
                
                $AppsFound = 0
                foreach ($App in $OfficeApps) {
                    $AppPath = Join-Path $OfficePath $App
                    if (Test-Path $AppPath) {
                        Write-DetectionLog "Found Office application: $AppPath"
                        
                        # Get version information
                        try {
                            $VersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($AppPath)
                            Write-DetectionLog "Application version: $($VersionInfo.FileVersion)"
                            $AppsFound++
                        }
                        catch {
                            Write-DetectionLog "Could not get version info for $AppPath" -Level Warning
                            $AppsFound++
                        }
                    }
                }
                
                # If we found at least 2 core Office apps, consider it installed
                if ($AppsFound -ge 2) {
                    Write-DetectionLog "Office installation confirmed - found $AppsFound core applications"
                    return $true
                }
            }
        }
        
        return $false
    }
    catch {
        Write-DetectionLog "Error checking Office executables: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to check Office Click-to-Run service
function Test-OfficeClickToRunService {
    try {
        Write-DetectionLog "Checking Office Click-to-Run service..."
        
        $Service = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
        if ($Service) {
            Write-DetectionLog "Click-to-Run service found - Status: $($Service.Status)"
            return $true
        }
        else {
            Write-DetectionLog "Click-to-Run service not found"
            return $false
        }
    }
    catch {
        Write-DetectionLog "Error checking Click-to-Run service: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Main detection logic
try {
    Write-DetectionLog "=== Microsoft 365 Apps Detection Started ==="
    Write-DetectionLog "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-DetectionLog "OS Version: $([System.Environment]::OSVersion.VersionString)"
    Write-DetectionLog "Computer Name: $($env:COMPUTERNAME)"
    Write-DetectionLog "User Context: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    
    # Perform detection checks
    $RegistryCheck = Test-OfficeClickToRunRegistry
    $ProductCheck = Test-OfficeInstalledProducts
    $ExecutableCheck = Test-OfficeExecutables
    $ServiceCheck = Test-OfficeClickToRunService
    
    Write-DetectionLog "Detection Results:"
    Write-DetectionLog "  Registry Check: $RegistryCheck"
    Write-DetectionLog "  Product Check: $ProductCheck"
    Write-DetectionLog "  Executable Check: $ExecutableCheck"
    Write-DetectionLog "  Service Check: $ServiceCheck"
    
    # Determine overall detection result
    # Office is considered installed if at least 2 of the 4 checks pass
    $PassedChecks = @($RegistryCheck, $ProductCheck, $ExecutableCheck, $ServiceCheck) | Where-Object { $_ -eq $true }
    $DetectionResult = $PassedChecks.Count -ge 2
    
    Write-DetectionLog "Passed Checks: $($PassedChecks.Count)/4"
    Write-DetectionLog "Overall Detection Result: $DetectionResult"
    
    if ($DetectionResult) {
        Write-DetectionLog "=== DETECTION SUCCESS: Microsoft 365 Apps is installed ===" -Level Info
        Write-Output "Microsoft 365 Apps detected successfully"
        exit 0
    }
    else {
        Write-DetectionLog "=== DETECTION FAILED: Microsoft 365 Apps is not installed ===" -Level Warning
        exit 1
    }
}
catch {
    Write-DetectionLog "=== DETECTION ERROR: $($_.Exception.Message) ===" -Level Error
    Write-DetectionLog "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}
finally {
    Write-DetectionLog "=== Microsoft 365 Apps Detection Completed ==="
}