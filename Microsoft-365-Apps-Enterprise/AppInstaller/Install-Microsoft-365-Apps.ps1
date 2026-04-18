<#
.SYNOPSIS
  Script to install/uninstall M365 Apps as a Win32 Apps.

.DESCRIPTION
    Script to install or uninstall Microsoft 365 Apps as a Win32 App during Autopilot by downloading the latest Office setup exe from evergreen url and running Setup.exe with provided configuration.xml or uninstall.xml file.

    OPTIMIZED FOR INTUNE/AUTOPILOT:
    - Supports both O365ProPlusRetail and O365BusinessRetail via parameter
    - Detects system architecture (AMD64 or ARM64).
    - Complete installation before exiting to ensure Microsoft 365 Apps are available at user login.
    - Proper logging to Intune Management Extension logs folder
    - Handles ClickToRunSvc to ensure clean exit for Intune detection
    - Cleanup after installation completes

.PARAMETER -Mode
    Specifies whether to Install or Uninstall Microsoft 365 Apps.

.PARAMETER -ProductID
    Specifies the Microsoft 365 Apps product to install or uninstall.
    Valid values: O365ProPlusRetail, O365BusinessRetail
    If not specified, uses the Product ID from the configuration.xml file.

.PARAMETER XMLUrl
    Optional URL to download configuration.xml from an external source.   

.EXAMPLE
    INTUNE CONFIGURATION:

    Install Microsoft 365 Apps for Enterprise:
    - Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install -ProductID O365ProPlusRetail
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall -ProductID O365ProPlusRetail
     
    Install Microsoft 365 Apps for Business:
    - Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install -ProductID O365BusinessRetail
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall -ProductID O365BusinessRetail

    Without -ProductID (uses XML default):
    - Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall

    Detection: 
    - Use Detect-Microsoft-365-Apps.ps1

.NOTES
    Version:        2.4
    Author:         UniFy-Endpoint
    Creation Date:  02-12-2025

#>


#region Parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("O365ProPlusRetail", "O365BusinessRetail")]
    [string]$ProductID,
    
    [Parameter(Mandatory = $false)]
    [string]$XMLUrl
)
#endregion Parameters

#region Variables
$SetupFolder = "$env:SystemRoot\Temp\OfficeSetup"
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFileName = "Microsoft-365-Apps-Setup.log"
$LogFilePath = Join-Path -Path $LogFolder -ChildPath $LogFileName
$SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
$ClickToRunServiceName = "ClickToRunSvc"
$OfficeC2RPath = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
$OfficeRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

# Cache timezone bias once at script start to avoid repeated CIM queries
$script:TimezoneBias = (Get-CimInstance -ClassName Win32_TimeZone | Select-Object -ExpandProperty Bias)
#endregion Variables

#region Functions
function Write-LogEntry {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = $LogFileName
    )
    
    # Ensure log folder exists
    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }
    
    # Construct time stamp for log entry using cached timezone bias
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", $script:TimezoneBias)
    $Date = (Get-Date -Format "MM-dd-yyyy")
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    
    # Construct final log entry (CMTrace compatible format)
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($FileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
    
    try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        if ($Severity -eq 1) { Write-Verbose -Message $Value }
        elseif ($Severity -eq 3) { Write-Warning -Message $Value }
    }
    catch {
        Write-Warning -Message "Unable to append log entry to $FileName file. Error: $($_.Exception.Message)"
    }
}

function Start-DownloadFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    
    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }

    $DestinationPath = Join-Path -Path $Path -ChildPath $Name
    
    # Use BITS for reliable download with fallback
    try {
        Write-LogEntry -Value "Downloading via BITS: $URL" -Severity 1
        Start-BitsTransfer -Source $URL -Destination $DestinationPath -ErrorAction Stop
        Write-LogEntry -Value "BITS download completed successfully" -Severity 1
    }
    catch {
        Write-LogEntry -Value "BITS failed, using WebClient: $($_.Exception.Message)" -Severity 2
        try {
            $WebClient = New-Object -TypeName System.Net.WebClient
            $WebClient.DownloadFile($URL, $DestinationPath)
            $WebClient.Dispose()
            Write-LogEntry -Value "WebClient download completed successfully" -Severity 1
        }
        catch {
            Write-LogEntry -Value "WebClient download failed: $($_.Exception.Message)" -Severity 3
            throw
        }
    }
}

function Invoke-FileCertVerification {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )
    
    $AuthSig = Get-AuthenticodeSignature -FilePath $FilePath
    $Cert = $AuthSig.SignerCertificate
    $CertStatus = $AuthSig.Status
    
    if ($Cert) {
        if ($Cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid") {
            $Chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
            $Chain.Build($Cert) | Out-Null
            $RootCert = $Chain.ChainElements | ForEach-Object { $_.Certificate } | Where-Object { $_.Subject -match "CN=Microsoft Root" }
            
            if (-not [string]::IsNullOrEmpty($RootCert)) {
                $TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $_.Thumbprint -eq $RootCert.Thumbprint }
                
                if (-not [string]::IsNullOrEmpty($TrustedRoot)) {
                    Write-LogEntry -Value "Verified setup file signed by: $($Cert.Issuer)" -Severity 1
                    return $true
                }
                else {
                    Write-LogEntry -Value "No trust found to root cert - aborting" -Severity 2
                    return $false
                }
            }
            else {
                Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2
                return $false
            }
        }
        else {
            Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2
            return $false
        }
    }
    else {
        Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
        return $false
    }
}

function Update-ConfigurationXmlProductID {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ProductID
    )
    
    try {
        [xml]$ConfigXml = Get-Content -Path $ConfigPath -Raw
        $ProductElement = $ConfigXml.Configuration.Add.Product
        
        if ($ProductElement) {
            $CurrentProductID = $ProductElement.ID
            if ($CurrentProductID -ne $ProductID) {
                Write-LogEntry -Value "Updating Product ID from '$CurrentProductID' to '$ProductID'" -Severity 1
                $ProductElement.ID = $ProductID
                $ConfigXml.Save($ConfigPath)
                Write-LogEntry -Value "Configuration XML Product ID updated successfully" -Severity 1
            }
            else {
                Write-LogEntry -Value "Product ID already set to '$ProductID'" -Severity 1
            }
        }
        return $true
    }
    catch {
        Write-LogEntry -Value "Failed to update Product ID in configuration XML: $($_.Exception.Message)" -Severity 3
        return $false
    }
}

function Update-UninstallXmlProductID {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ProductID
    )
    
    try {
        [xml]$ConfigXml = Get-Content -Path $ConfigPath -Raw
        $ProductElement = $ConfigXml.Configuration.Remove.Product
        
        if ($ProductElement) {
            $CurrentProductID = $ProductElement.ID
            if ($CurrentProductID -ne $ProductID) {
                Write-LogEntry -Value "Updating Uninstall Product ID from '$CurrentProductID' to '$ProductID'" -Severity 1
                $ProductElement.ID = $ProductID
                $ConfigXml.Save($ConfigPath)
                Write-LogEntry -Value "Uninstall XML Product ID updated successfully" -Severity 1
            }
            else {
                Write-LogEntry -Value "Uninstall Product ID already set to '$ProductID'" -Severity 1
            }
        }
        return $true
    }
    catch {
        Write-LogEntry -Value "Failed to update Product ID in uninstall XML: $($_.Exception.Message)" -Severity 3
        return $false
    }
}

function Test-M365AppsInstalled {
    # Check if Office Click-to-Run executable exists
    if (Test-Path -Path $OfficeC2RPath) {
        # Verify Office registry key exists with version
        if (Test-Path -Path $OfficeRegPath) {
            $VersionToReport = Get-ItemProperty -Path $OfficeRegPath -Name "VersionToReport" -ErrorAction SilentlyContinue
            if ($VersionToReport.VersionToReport) {
                return $true
            }
        }
    }
    return $false
}

function Test-ProductInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductID
    )
    
    try {
        $ProductReleaseIds = (Get-ItemProperty -Path $OfficeRegPath -Name "ProductReleaseIds" -ErrorAction SilentlyContinue).ProductReleaseIds
        if ($ProductReleaseIds) {
            $InstalledProducts = $ProductReleaseIds -split ","
            return $InstalledProducts -contains $ProductID
        }
        return $false
    }
    catch {
        return $false
    }
}

function Get-InstalledOfficeProducts {
    try {
        $ProductReleaseIds = (Get-ItemProperty -Path $OfficeRegPath -Name "ProductReleaseIds" -ErrorAction SilentlyContinue).ProductReleaseIds
        if ($ProductReleaseIds) {
            return $ProductReleaseIds
        }
        return $null
    }
    catch {
        return $null
    }
}

function Get-OfficeInstallationStatus {
    <#
    .SYNOPSIS
        Checks the Office Click-to-Run installation status from registry.
    .DESCRIPTION
        Returns the current installation state:
        - "Complete" - Installation finished successfully
        - "Installing" - Installation in progress
        - "NotFound" - Office not detected
    #>
    
    try {
        # Check ExecutingScenario - if empty or null, no active operation
        $ExecutingScenario = Get-ItemProperty -Path $OfficeRegPath -Name "ExecutingScenario" -ErrorAction SilentlyContinue
        
        # Check for version to confirm installation
        $VersionToReport = Get-ItemProperty -Path $OfficeRegPath -Name "VersionToReport" -ErrorAction SilentlyContinue
        
        if ($VersionToReport.VersionToReport) {
            # Version exists, check if still executing
            if ([string]::IsNullOrEmpty($ExecutingScenario.ExecutingScenario)) {
                return "Complete"
            }
            else {
                Write-LogEntry -Value "Office operation in progress: $($ExecutingScenario.ExecutingScenario)" -Severity 1
                return "Installing"
            }
        }
        else {
            return "NotFound"
        }
    }
    catch {
        return "NotFound"
    }
}

function Wait-OfficeInstallationComplete {
    param(
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 120,
        
        [Parameter(Mandatory = $false)]
        [int]$CheckIntervalSeconds = 5
    )
    
    Write-LogEntry -Value "Verifying Office installation completion (timeout: $TimeoutSeconds seconds)..." -Severity 1
    
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    while ($Stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $Status = Get-OfficeInstallationStatus
        
        switch ($Status) {
            "Complete" {
                $VersionToReport = (Get-ItemProperty -Path $OfficeRegPath -Name "VersionToReport" -ErrorAction SilentlyContinue).VersionToReport
                Write-LogEntry -Value "Office installation verified complete. Version: $VersionToReport" -Severity 1
                $Stopwatch.Stop()
                return $true
            }
            "Installing" {
                Write-LogEntry -Value "Office still configuring, waiting... (Elapsed: $([math]::Round($Stopwatch.Elapsed.TotalSeconds))s)" -Severity 1
            }
            "NotFound" {
                Write-LogEntry -Value "Office registry not found yet, waiting... (Elapsed: $([math]::Round($Stopwatch.Elapsed.TotalSeconds))s)" -Severity 1
            }
        }
        
        Start-Sleep -Seconds $CheckIntervalSeconds
    }
    
    $Stopwatch.Stop()
    Write-LogEntry -Value "Timeout reached waiting for Office installation verification" -Severity 2
    
    # Final check - if version exists, consider it successful
    if (Test-M365AppsInstalled) {
        Write-LogEntry -Value "Office installation detected despite timeout - proceeding" -Severity 1
        return $true
    }
    
    return $false
}

function Stop-ClickToRunService {
    param(
        [Parameter(Mandatory = $false)]
        [int]$WaitSeconds = 30
    )
    
    Write-LogEntry -Value "Stopping Click-to-Run service for clean Intune detection..." -Severity 1
    
    try {
        $Service = Get-Service -Name $ClickToRunServiceName -ErrorAction SilentlyContinue
        
        if ($Service) {
            if ($Service.Status -eq "Running") {
                Write-LogEntry -Value "Stopping $ClickToRunServiceName service..." -Severity 1
                
                # Stop the service gracefully
                Stop-Service -Name $ClickToRunServiceName -Force -ErrorAction Stop
                
                # Wait for service to stop
                $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                while ($Stopwatch.Elapsed.TotalSeconds -lt $WaitSeconds) {
                    $Service = Get-Service -Name $ClickToRunServiceName -ErrorAction SilentlyContinue
                    if ($Service.Status -eq "Stopped") {
                        Write-LogEntry -Value "$ClickToRunServiceName service stopped successfully" -Severity 1
                        $Stopwatch.Stop()
                        return $true
                    }
                    Start-Sleep -Seconds 2
                }
                $Stopwatch.Stop()
                
                Write-LogEntry -Value "Timeout waiting for $ClickToRunServiceName to stop" -Severity 2
                return $false
            }
            else {
                Write-LogEntry -Value "$ClickToRunServiceName service is not running (Status: $($Service.Status))" -Severity 1
                return $true
            }
        }
        else {
            Write-LogEntry -Value "$ClickToRunServiceName service not found" -Severity 2
            return $true
        }
    }
    catch {
        Write-LogEntry -Value "Error stopping $ClickToRunServiceName service: $($_.Exception.Message)" -Severity 2
        return $false
    }
}

function Invoke-OfficeCleanup {
    Write-LogEntry -Value "Starting cleanup of setup folder" -Severity 1
    if (Test-Path -Path $SetupFolder) {
        try {
            Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Stop
            Write-LogEntry -Value "Setup folder cleaned up successfully" -Severity 1
        }
        catch {
            Write-LogEntry -Value "Warning: Could not remove setup folder: $($_.Exception.Message)" -Severity 2
        }
    }
}
#endregion Functions

#region Main Script
try {
    # Initialize logging
    Write-LogEntry -Value "========================================" -Severity 1
    Write-LogEntry -Value "Starting Microsoft 365 Apps $Mode process" -Severity 1
    Write-LogEntry -Value "Script version: 2.4" -Severity 1
    Write-LogEntry -Value "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Severity 1
    Write-LogEntry -Value "System architecture: $($env:PROCESSOR_ARCHITECTURE)" -Severity 1
    if ($ProductID) {
        Write-LogEntry -Value "Product ID specified: $ProductID" -Severity 1
    }
    else {
        Write-LogEntry -Value "Product ID not specified - using value from XML file" -Severity 1
    }
    
    # Log currently installed Office products
    $CurrentProducts = Get-InstalledOfficeProducts
    if ($CurrentProducts) {
        Write-LogEntry -Value "Currently installed Office products: $CurrentProducts" -Severity 1
    }
    else {
        Write-LogEntry -Value "No Office products currently detected" -Severity 1
    }
    
    # Cleanup any existing setup folder
    Invoke-OfficeCleanup
    
    # Create setup folder
    Write-LogEntry -Value "Creating setup folder: $SetupFolder" -Severity 1
    if (-not (Test-Path -Path $SetupFolder)) {
        New-Item -Path $SetupFolder -ItemType Directory -Force | Out-Null
    }
    
    # Download Office setup.exe
    Write-LogEntry -Value "Downloading Office setup executable from: $SetupEverGreenURL" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    
    # Verify download
    if (-not (Test-Path -Path $SetupFilePath)) {
        Write-LogEntry -Value "Error: Setup file not found after download" -Severity 3
        Invoke-OfficeCleanup
        exit 1
    }
    
    Write-LogEntry -Value "Setup file ready at: $SetupFilePath" -Severity 1
    
    # Get Office version info
    $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion
    Write-LogEntry -Value "Office C2R Setup version: $OfficeCR2Version" -Severity 1
    
    # Verify certificate
    if (-not (Invoke-FileCertVerification -FilePath $SetupFilePath)) {
        Write-LogEntry -Value "Error: Unable to verify setup file signature - aborting" -Severity 3
        Invoke-OfficeCleanup
        exit 1
    }
    
    # Initialize exit code
    $ExitCode = $null
    
    # Handle Install or Uninstall mode
    switch ($Mode) {
        "Install" {
            if ($XMLUrl) {
                Write-LogEntry -Value "Downloading configuration.xml from: $XMLUrl" -Severity 1
                try {
                    Start-DownloadFile -URL $XMLUrl -Path $SetupFolder -Name "configuration.xml"
                    Write-LogEntry -Value "Configuration.xml downloaded successfully" -Severity 1
                }
                catch {
                    Write-LogEntry -Value "Failed to download configuration.xml: $($_.Exception.Message)" -Severity 3
                    Invoke-OfficeCleanup
                    exit 1
                }
            }
            else {
                $LocalConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "configuration.xml"
                if (-not (Test-Path -Path $LocalConfigPath)) {
                    Write-LogEntry -Value "Error: Local configuration.xml not found at: $LocalConfigPath" -Severity 3
                    Invoke-OfficeCleanup
                    exit 1
                }
                Write-LogEntry -Value "Using local configuration.xml" -Severity 1
                Copy-Item -Path $LocalConfigPath -Destination $SetupFolder -Force -ErrorAction Stop
            }
            
            $ConfigFilePath = Join-Path -Path $SetupFolder -ChildPath "configuration.xml"
            
            # Update Product ID if specified
            if ($ProductID) {
                if (-not (Update-ConfigurationXmlProductID -ConfigPath $ConfigFilePath -ProductID $ProductID)) {
                    Write-LogEntry -Value "Warning: Could not update Product ID in configuration file" -Severity 2
                }
            }
            
            Write-LogEntry -Value "Starting Microsoft 365 Apps installation..." -Severity 1
            $Arguments = "/configure `"$ConfigFilePath`""
            Write-LogEntry -Value "Running: $SetupFilePath $Arguments" -Severity 1
            
            $Process = Start-Process -FilePath $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
            $ExitCode = $Process.ExitCode
            
            Write-LogEntry -Value "Office setup.exe completed with exit code: $ExitCode" -Severity 1
            
            # Post-installation verification
            if ($ExitCode -eq 0) {
                # Wait for installation to fully complete using registry check
                $InstallVerified = Wait-OfficeInstallationComplete -TimeoutSeconds 120 -CheckIntervalSeconds 5
                
                if ($InstallVerified) {
                    # Log installed products after installation
                    $PostInstallProducts = Get-InstalledOfficeProducts
                    Write-LogEntry -Value "Installed Office products after installation: $PostInstallProducts" -Severity 1
                    
                    # Stop ClickToRunSvc to ensure clean exit for Intune detection
                    $ServiceStopped = Stop-ClickToRunService -WaitSeconds 30
                    
                    if (-not $ServiceStopped) {
                        Write-LogEntry -Value "Warning: Could not stop ClickToRunSvc, but installation completed" -Severity 2
                    }
                }
                else {
                    Write-LogEntry -Value "Warning: Could not verify installation completion, but setup.exe returned success" -Severity 2
                }
            }
        }
        
        "Uninstall" {
            # Determine target ProductID
            $TargetProductID = if ($ProductID) { $ProductID } else { "O365ProPlusRetail" }
            Write-LogEntry -Value "Target product to uninstall: $TargetProductID" -Severity 1
            
            # Check if target product is installed
            if (-not (Test-ProductInstalled -ProductID $TargetProductID)) {
                Write-LogEntry -Value "$TargetProductID not detected - nothing to uninstall" -Severity 1
                $ExitCode = 0
            }
            else {
                Write-LogEntry -Value "$TargetProductID is installed - proceeding with uninstall" -Severity 1
                
                $LocalUninstallPath = Join-Path -Path $PSScriptRoot -ChildPath "uninstall.xml"
                if (-not (Test-Path -Path $LocalUninstallPath)) {
                    Write-LogEntry -Value "Error: Local uninstall.xml not found at: $LocalUninstallPath" -Severity 3
                    Invoke-OfficeCleanup
                    exit 1
                }
                
                Write-LogEntry -Value "Using local uninstall.xml" -Severity 1
                Copy-Item -Path $LocalUninstallPath -Destination $SetupFolder -Force -ErrorAction Stop
                
                $UninstallConfigPath = Join-Path -Path $SetupFolder -ChildPath "uninstall.xml"
                
                # Update Product ID if specified
                if ($ProductID) {
                    if (-not (Update-UninstallXmlProductID -ConfigPath $UninstallConfigPath -ProductID $ProductID)) {
                        Write-LogEntry -Value "Warning: Could not update Product ID in uninstall file" -Severity 2
                    }
                }
                
                Write-LogEntry -Value "Starting Microsoft 365 Apps uninstallation..." -Severity 1
                $Arguments = "/configure `"$UninstallConfigPath`""
                Write-LogEntry -Value "Running: $SetupFilePath $Arguments" -Severity 1
                
                $Process = Start-Process -FilePath $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
                $ExitCode = $Process.ExitCode
                
                Write-LogEntry -Value "Office uninstall completed with exit code: $ExitCode" -Severity 1
                
                # Verify uninstall result
                if ($ExitCode -eq 0 -or $ExitCode -eq -1) {
                    # Brief wait for registry to update
                    Write-LogEntry -Value "Waiting for registry to update..." -Severity 1
                    Start-Sleep -Seconds 5
                    
                    # Check if target product was actually removed
                    if (-not (Test-ProductInstalled -ProductID $TargetProductID)) {
                        Write-LogEntry -Value "$TargetProductID successfully removed" -Severity 1
                        
                        # Log remaining products
                        $RemainingProducts = Get-InstalledOfficeProducts
                        if ($RemainingProducts) {
                            Write-LogEntry -Value "Remaining installed Office products: $RemainingProducts" -Severity 1
                        }
                        else {
                            Write-LogEntry -Value "No Office products remaining" -Severity 1
                        }
                        
                        $ExitCode = 0
                    }
                    else {
                        Write-LogEntry -Value "$TargetProductID still present after uninstall attempt" -Severity 3
                        $ExitCode = 1
                    }
                }
            }
        }
    }
    
    # Cleanup
    Invoke-OfficeCleanup
    
    # Final status
    if ($ExitCode -eq 0) {
        Write-LogEntry -Value "Microsoft 365 Apps $Mode completed successfully" -Severity 1
        Write-LogEntry -Value "========================================" -Severity 1
        exit 0
    }
    else {
        Write-LogEntry -Value "Microsoft 365 Apps $Mode failed with exit code: $ExitCode" -Severity 3
        Write-LogEntry -Value "========================================" -Severity 1
        exit $ExitCode
    }
}
catch {
    Write-LogEntry -Value "Critical error during $Mode process: $($_.Exception.Message)" -Severity 3
    Write-LogEntry -Value "Stack trace: $($_.ScriptStackTrace)" -Severity 3
    Invoke-OfficeCleanup
    exit 1
}
#endregion Main Script