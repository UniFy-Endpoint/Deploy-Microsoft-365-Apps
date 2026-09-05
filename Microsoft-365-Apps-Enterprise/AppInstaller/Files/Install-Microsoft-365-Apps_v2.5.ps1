<#
.SYNOPSIS
  Script to install/uninstall Microsoft 365 Apps as an Intune Win32 App.

.DESCRIPTION
    Installs or uninstalls Microsoft 365 Apps by downloading the latest Office Deployment Tool
    setup.exe from the evergreen URL and running it against configuration.xml / uninstall.xml.

    OPTIMIZED FOR INTUNE/AUTOPILOT:
    - Supports both O365ProPlusRetail and O365BusinessRetail via parameter
    - Runs the full installation before exiting, so Office is present at first user logon
    - Logs to the Intune Management Extension logs folder in CMTrace format
    - Stops ClickToRunSvc after a successful install for a clean Intune detection pass
    - Removes only the targeted product, leaving other Office versions untouched
    - Cleans up the setup folder afterwards

    CHANGES IN 2.5 (vs 2.4):
    - Re-launches itself in 64-bit PowerShell when started from a 32-bit host. The Intune
      Management Extension is a 32-bit process, so setup ran under WOW64: $env:ProgramFiles
      resolved to "Program Files (x86)" and HKLM\SOFTWARE\Microsoft\Office was redirected to
      WOW6432Node. Every post-install verification therefore failed, Wait-OfficeInstallation-
      Complete burned its full timeout, and Stop-ClickToRunService was never reached.
    - All registry reads use the 64-bit view explicitly, so results no longer depend on host
      bitness even if the relaunch is bypassed.
    - Waits for the Click-to-Run engine to go idle BEFORE running setup.exe. Running an
      uninstall immediately after an install previously failed until the device was rebooted,
      because the engine was still executing the install scenario.
    - Reports which languages ended up installed, so language-pack problems are visible in
      the log instead of only surfacing as a failed detection.
    - Cleanup no longer deletes the setup folder while an Office process still has it open.

.PARAMETER Mode
    Install or Uninstall.

.PARAMETER ProductID
    O365ProPlusRetail or O365BusinessRetail. When omitted, the Product ID inside the XML file
    is used unchanged.

.PARAMETER XMLUrl
    Optional URL to download configuration.xml from an external source instead of using the
    local copy next to this script.

.PARAMETER RequiredLanguages
    Languages that configuration.xml is expected to install. Reported in the log after a
    successful install. Does not fail the install - the detection script is the gate.

.EXAMPLE
    INTUNE - WIN32 SCRIPT PACKAGE (Microsoft Win32 Content Prep Tool)

    Install Microsoft 365 Apps for business:
    - Install command:   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Install -ProductID O365BusinessRetail
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Uninstall -ProductID O365BusinessRetail

    Install Microsoft 365 Apps for enterprise:
    - Install command:   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Install -ProductID O365ProPlusRetail
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Uninstall -ProductID O365ProPlusRetail

    Without -ProductID (uses the Product ID already in the XML):
    - Install command:   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Install
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Uninstall

.EXAMPLE
    INTUNE - PSAPPDEPLOYTOOLKIT PACKAGE (..\PSADT-Package)

    This script is wrapped by Invoke-AppDeployToolkit.ps1, which calls it with the same
    parameters. Use these commands in Intune instead of calling PowerShell directly:

    - Install command:   Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Auto
    - Uninstall command: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent

    Append -ProductID O365BusinessRetail to both commands to reuse the package for business.

.EXAMPLE
    Detection (both packaging methods):
    - Use Detection\Detect-Microsoft-365-Apps_v2.0.ps1 as a custom detection script.

.NOTES
    Version:        2.5
    Author:         UniFy-Endpoint
    Creation Date:  02-12-2025
    Updated:        04-09-2026

    Exit codes:
    - 0     Success
    - 1     Script-level failure (download, signature, missing XML, unhandled error)
    - other Passed through unchanged from setup.exe
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
    [string]$XMLUrl,

    [Parameter(Mandatory = $false)]
    [string[]]$RequiredLanguages = @("nl-nl", "en-us")
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
$ScriptVersion = "2.5"

# Registry paths are relative to the 64-bit HKLM hive, opened explicitly below.
$C2RConfigKey = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$C2RProductKey = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
$UninstallKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

# $env:ProgramFiles is redirected inside a 32-bit host; ProgramW6432 never is.
$ProgramFiles64 = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
$OfficeC2RPath = Join-Path -Path $ProgramFiles64 -ChildPath "Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"

# Processes that indicate the Click-to-Run engine is mid-scenario.
$OfficeSetupProcesses = @("setup", "OfficeC2RClient")

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

function Get-HKLM64 {
    # Always read the 64-bit hive, regardless of the bitness of the host process.
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        [Microsoft.Win32.RegistryView]::Registry64)
}

function Get-RegistryValue64 {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $Base = $null
    $SubKey = $null
    try {
        $Base = Get-HKLM64
        $SubKey = $Base.OpenSubKey($Key)
        if ($null -eq $SubKey) { return $null }
        return $SubKey.GetValue($Name)
    }
    catch {
        return $null
    }
    finally {
        if ($SubKey) { $SubKey.Close() }
        if ($Base) { $Base.Close() }
    }
}

function Test-RegistryKey64 {
    param(
        [Parameter(Mandatory = $true)][string]$Key
    )

    $Base = $null
    $SubKey = $null
    try {
        $Base = Get-HKLM64
        $SubKey = $Base.OpenSubKey($Key)
        return ($null -ne $SubKey)
    }
    catch {
        return $false
    }
    finally {
        if ($SubKey) { $SubKey.Close() }
        if ($Base) { $Base.Close() }
    }
}

function Get-RegistrySubKeyNames64 {
    param(
        [Parameter(Mandatory = $true)][string]$Key
    )

    $Base = $null
    $SubKey = $null
    try {
        $Base = Get-HKLM64
        $SubKey = $Base.OpenSubKey($Key)
        if ($null -eq $SubKey) { return @() }
        return @($SubKey.GetSubKeyNames())
    }
    catch {
        return @()
    }
    finally {
        if ($SubKey) { $SubKey.Close() }
        if ($Base) { $Base.Close() }
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
    if (Test-Path -Path $OfficeC2RPath) {
        $VersionToReport = Get-RegistryValue64 -Key $C2RConfigKey -Name "VersionToReport"
        if (-not [string]::IsNullOrWhiteSpace($VersionToReport)) {
            return $true
        }
    }
    return $false
}

function Get-InstalledOfficeProducts {
    $ProductReleaseIds = Get-RegistryValue64 -Key $C2RConfigKey -Name "ProductReleaseIds"
    if ([string]::IsNullOrWhiteSpace($ProductReleaseIds)) { return $null }
    return $ProductReleaseIds
}

function Test-ProductInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductID
    )

    $ProductReleaseIds = Get-InstalledOfficeProducts
    if ([string]::IsNullOrWhiteSpace($ProductReleaseIds)) { return $false }

    $InstalledProducts = @($ProductReleaseIds -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    return ($InstalledProducts -contains $ProductID)
}

function Get-InstalledOfficeLanguages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductID
    )

    $Languages = New-Object System.Collections.Generic.List[string]

    # Per-install culture subkeys, e.g. ProductReleaseIDs\<InstallID>\O365ProPlusRetail.16\nl-nl
    foreach ($InstallID in Get-RegistrySubKeyNames64 -Key $C2RProductKey) {
        foreach ($Culture in Get-RegistrySubKeyNames64 -Key "$C2RProductKey\$InstallID\$ProductID.16") {
            if ($Culture -ne "x-none" -and -not $Languages.Contains($Culture)) {
                $Languages.Add($Culture)
            }
        }
    }

    return $Languages.ToArray()
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
        $ExecutingScenario = Get-RegistryValue64 -Key $C2RConfigKey -Name "ExecutingScenario"
        $VersionToReport = Get-RegistryValue64 -Key $C2RConfigKey -Name "VersionToReport"

        if (-not [string]::IsNullOrWhiteSpace($VersionToReport)) {
            if ([string]::IsNullOrWhiteSpace($ExecutingScenario)) {
                return "Complete"
            }
            else {
                Write-LogEntry -Value "Office operation in progress: $ExecutingScenario" -Severity 1
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

function Wait-OfficeEngineIdle {
    <#
    .SYNOPSIS
        Waits for the Click-to-Run engine to finish whatever it is doing.
    .DESCRIPTION
        Running setup.exe while a previous scenario is still executing is the reason an
        uninstall issued immediately after an install used to do nothing until the device
        was rebooted. Waiting for ExecutingScenario to clear and for the setup processes to
        exit makes the uninstall take effect without a restart.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 900,

        [Parameter(Mandatory = $false)]
        [int]$CheckIntervalSeconds = 10
    )

    Write-LogEntry -Value "Waiting for the Click-to-Run engine to become idle (timeout: $TimeoutSeconds seconds)..." -Severity 1

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($Stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $ExecutingScenario = Get-RegistryValue64 -Key $C2RConfigKey -Name "ExecutingScenario"
        $ActiveProcesses = @(Get-Process -Name $OfficeSetupProcesses -ErrorAction SilentlyContinue)

        if ([string]::IsNullOrWhiteSpace($ExecutingScenario) -and $ActiveProcesses.Count -eq 0) {
            Write-LogEntry -Value "Click-to-Run engine is idle after $([math]::Round($Stopwatch.Elapsed.TotalSeconds))s" -Severity 1
            $Stopwatch.Stop()
            return $true
        }

        if (-not [string]::IsNullOrWhiteSpace($ExecutingScenario)) {
            Write-LogEntry -Value "Engine busy with scenario '$ExecutingScenario', waiting... (Elapsed: $([math]::Round($Stopwatch.Elapsed.TotalSeconds))s)" -Severity 1
        }
        else {
            Write-LogEntry -Value "Setup processes still running: $($ActiveProcesses.Name -join ', '). Waiting... (Elapsed: $([math]::Round($Stopwatch.Elapsed.TotalSeconds))s)" -Severity 1
        }

        Start-Sleep -Seconds $CheckIntervalSeconds
    }

    $Stopwatch.Stop()
    Write-LogEntry -Value "Timeout waiting for the Click-to-Run engine to become idle - continuing anyway" -Severity 2
    return $false
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
                $VersionToReport = Get-RegistryValue64 -Key $C2RConfigKey -Name "VersionToReport"
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

                Stop-Service -Name $ClickToRunServiceName -Force -ErrorAction Stop

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
    if (-not (Test-Path -Path $SetupFolder)) { return }

    # Deleting the folder while setup.exe still has it open leaves a partial tree behind.
    $ActiveProcesses = @(Get-Process -Name $OfficeSetupProcesses -ErrorAction SilentlyContinue)
    if ($ActiveProcesses.Count -gt 0) {
        Write-LogEntry -Value "Office setup processes still running ($($ActiveProcesses.Name -join ', ')) - skipping cleanup" -Severity 2
        return
    }

    Write-LogEntry -Value "Starting cleanup of setup folder" -Severity 1
    try {
        Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Stop
        Write-LogEntry -Value "Setup folder cleaned up successfully" -Severity 1
    }
    catch {
        Write-LogEntry -Value "Warning: Could not remove setup folder: $($_.Exception.Message)" -Severity 2
    }
}

function Invoke-OfficeSetup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    # Never start setup.exe while the engine is mid-scenario.
    Wait-OfficeEngineIdle -TimeoutSeconds 900 -CheckIntervalSeconds 10 | Out-Null

    $Arguments = "/configure `"$ConfigPath`""
    Write-LogEntry -Value "Running: $SetupFilePath $Arguments" -Severity 1

    $Process = Start-Process -FilePath $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
    return $Process.ExitCode
}
#endregion Functions

#region Main Script
try {
    # The Intune Management Extension is 32-bit, so without this the whole script runs under
    # WOW64 and every Office path and registry read is silently redirected.
    if ([System.Environment]::Is64BitOperatingSystem -and -not [System.Environment]::Is64BitProcess) {
        $NativePowerShell = Join-Path -Path $env:SystemRoot -ChildPath "SysNative\WindowsPowerShell\v1.0\powershell.exe"

        if (Test-Path -Path $NativePowerShell) {
            Write-LogEntry -Value "Running in a 32-bit host - relaunching in 64-bit PowerShell" -Severity 1

            $RelaunchArguments = @(
                "-NoProfile"
                "-NonInteractive"
                "-ExecutionPolicy", "Bypass"
                "-File", "`"$PSCommandPath`""
                "-Mode", $Mode
            )
            if ($ProductID) { $RelaunchArguments += @("-ProductID", $ProductID) }
            if ($XMLUrl) { $RelaunchArguments += @("-XMLUrl", "`"$XMLUrl`"") }
            if ($RequiredLanguages -and $RequiredLanguages.Count -gt 0) {
                $RelaunchArguments += @("-RequiredLanguages", ($RequiredLanguages -join ","))
            }

            $Relaunched = Start-Process -FilePath $NativePowerShell -ArgumentList $RelaunchArguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
            Write-LogEntry -Value "64-bit instance completed with exit code: $($Relaunched.ExitCode)" -Severity 1
            exit $Relaunched.ExitCode
        }

        Write-LogEntry -Value "SysNative not available - continuing in the 32-bit host" -Severity 2
    }

    # Initialize logging
    Write-LogEntry -Value "========================================" -Severity 1
    Write-LogEntry -Value "Starting Microsoft 365 Apps $Mode process" -Severity 1
    Write-LogEntry -Value "Script version: $ScriptVersion" -Severity 1
    Write-LogEntry -Value "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Severity 1
    Write-LogEntry -Value "System architecture: $($env:PROCESSOR_ARCHITECTURE), 64-bit process: $([System.Environment]::Is64BitProcess)" -Severity 1
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
            $ExitCode = Invoke-OfficeSetup -ConfigPath $ConfigFilePath

            Write-LogEntry -Value "Office setup.exe completed with exit code: $ExitCode" -Severity 1

            # Post-installation verification
            if ($ExitCode -eq 0) {
                $InstallVerified = Wait-OfficeInstallationComplete -TimeoutSeconds 120 -CheckIntervalSeconds 5

                if ($InstallVerified) {
                    $PostInstallProducts = Get-InstalledOfficeProducts
                    Write-LogEntry -Value "Installed Office products after installation: $PostInstallProducts" -Severity 1

                    # Report the language state so a missing language pack is visible in the
                    # log rather than only showing up later as a failed detection.
                    $EffectiveProductID = if ($ProductID) { $ProductID } else { @($PostInstallProducts -split "," | ForEach-Object { $_.Trim() })[0] }
                    if ($EffectiveProductID) {
                        $InstalledLanguages = Get-InstalledOfficeLanguages -ProductID $EffectiveProductID
                        Write-LogEntry -Value "Languages installed for ${EffectiveProductID}: $($InstalledLanguages -join ', ')" -Severity 1

                        foreach ($Language in $RequiredLanguages) {
                            if ($InstalledLanguages -contains $Language) {
                                Write-LogEntry -Value "Required language '$Language' is installed" -Severity 1
                            }
                            else {
                                Write-LogEntry -Value "Required language '$Language' is NOT installed - detection will fail and Intune will retry" -Severity 3
                            }
                        }
                    }

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
                $ExitCode = Invoke-OfficeSetup -ConfigPath $UninstallConfigPath

                Write-LogEntry -Value "Office uninstall completed with exit code: $ExitCode" -Severity 1

                # Verify uninstall result
                if ($ExitCode -eq 0 -or $ExitCode -eq -1) {
                    # Let the engine finish removing the product before checking the registry.
                    Wait-OfficeEngineIdle -TimeoutSeconds 600 -CheckIntervalSeconds 10 | Out-Null
                    Start-Sleep -Seconds 5

                    if (-not (Test-ProductInstalled -ProductID $TargetProductID)) {
                        Write-LogEntry -Value "$TargetProductID successfully removed" -Severity 1

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

