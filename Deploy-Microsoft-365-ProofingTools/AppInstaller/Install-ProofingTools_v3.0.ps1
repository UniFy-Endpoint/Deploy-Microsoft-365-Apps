<#
.SYNOPSIS
  Script to install/uninstall Microsoft 365 Apps proofing tools as an Intune Win32 App.

.DESCRIPTION
    Downloads the latest Office Deployment Tool setup.exe from the evergreen URL and runs it
    against install.xml / uninstall.xml, with the requested language written into the XML.

    CHANGES IN 3.0 (vs 2.1):
    - FIXED: the install path called Test-ProofingToolsInstallation, but the function is named
      Test-ProofingToolsStatus. The call raised CommandNotFoundException, which the surrounding
      catch turned into ExitCode 1. Every install therefore reported FAILURE to Intune even
      when the proofing tools installed correctly. Uninstall was unaffected because it never
      called the verification function at all.
    - Verification is now proofing-specific. 2.1 treated "the language appears in
      InstalledLanguages" as proof that the proofing tools were installed, but a plain language
      pack also puts the language there - so the proofing tools app could verify as installed
      when only the language pack was present. That value is also absent on current
      Click-to-Run builds, so in practice the check fell through to DisplayName matching.
      3.0 only accepts the ProofingTools culture key or a "<Product> - <lang>.proof" ARP entry.
    - Re-launches itself in 64-bit PowerShell when started from the 32-bit Intune Management
      Extension host, and reads the registry through the explicit 64-bit view.
    - Waits for the Click-to-Run engine to go idle before running setup.exe, so an uninstall
      issued straight after an install takes effect without a reboot.
    - Uninstall is now verified the same way as install.
    - Logs to the Intune Management Extension logs folder, matching the other packages in this
      solution (2.1 logged to %SystemRoot%\Temp).

.PARAMETER LanguageID
    Language in ODT format, e.g. nl-nl or en-us.

.PARAMETER Mode
    Install or Uninstall.

.PARAMETER MaxWaitMinutes
    How long to wait for the Click-to-Run engine to finish applying the change.

.EXAMPLE
    INTUNE - WIN32 SCRIPT PACKAGE (Microsoft Win32 Content Prep Tool)

    - Install command:   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-ProofingTools_v3.0.ps1 -LanguageID "nl-nl" -Mode Install
    - Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-ProofingTools_v3.0.ps1 -LanguageID "nl-nl" -Mode Uninstall

.EXAMPLE
    INTUNE - PSAPPDEPLOYTOOLKIT PACKAGE

    This script is wrapped by Invoke-AppDeployToolkit.ps1, which calls it with the same
    parameters. Use these commands in Intune instead of calling PowerShell directly:

    - Install command:   Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Auto
    - Uninstall command: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent

.EXAMPLE
    Detection (both packaging methods):
    - Use Detection\Detect-ProofingTools_v3.0.ps1 as a custom detection script.

.NOTES
  Version:    3.0
  Author:     UniFy-Endpoint
  Updated:    04-09-2026

  Exit codes:
  - 0     Success
  - 1     Failure (download, signature, setup.exe failure, or verification timeout)
#>

#region Parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LanguageID,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode,

    [Parameter(Mandatory = $false)]
    [int]$MaxWaitMinutes = 30
)
#endregion Parameters

#region Variables
$SetupFolder = "$env:SystemRoot\Temp\OfficeSetup"
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFileName = "M365ProofingToolsSetup.log"
$LogFilePath = Join-Path -Path $LogFolder -ChildPath $LogFileName
$SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
$ScriptVersion = "3.0"

# Registry paths are relative to the 64-bit HKLM hive, opened explicitly below.
$C2RConfigKey = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$C2RProductKey = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
$UninstallKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

# Base Click-to-Run products the proofing tools can attach to.
$BaseProductIDs = @("O365ProPlusRetail", "O365BusinessRetail")

# The Click-to-Run product ID that ODT uses for proofing tools.
$ProofingProductID = "ProofingTools"

# Processes that indicate the Click-to-Run engine is mid-scenario.
$OfficeSetupProcesses = @("setup", "OfficeC2RClient")

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

    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", $script:TimezoneBias)
    $Date = (Get-Date -Format "MM-dd-yyyy")
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
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

    try {
        Write-LogEntry -Value "Downloading via BITS: $URL" -Severity 1
        Start-BitsTransfer -Source $URL -Destination $DestinationPath -ErrorAction Stop
        Write-LogEntry -Value "BITS download completed successfully" -Severity 1
    }
    catch {
        Write-LogEntry -Value "BITS failed, using WebClient: $($_.Exception.Message)" -Severity 2
        $WebClient = New-Object -TypeName System.Net.WebClient
        $WebClient.DownloadFile($URL, $DestinationPath)
        $WebClient.Dispose()
        Write-LogEntry -Value "WebClient download completed successfully" -Severity 1
    }
}

function Invoke-XMLUpdate {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$LanguageID,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$Mode
    )

    [xml]$XmlDoc = Get-Content -Path $FileName -Raw

    if ($Mode -eq "Install") {
        $XmlDoc.Configuration.Add.Product.Language.ID = $LanguageID
    }
    else {
        $XmlDoc.Configuration.Remove.Product.Language.ID = $LanguageID
    }

    $XmlDoc.Save($FileName)
    Write-LogEntry -Value "Set language '$LanguageID' in $(Split-Path -Path $FileName -Leaf)" -Severity 1
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

    if (-not $Cert) {
        Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
        return $false
    }

    if ($Cert.Subject -notmatch "O=Microsoft Corporation" -or $CertStatus -ne "Valid") {
        Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2
        return $false
    }

    $Chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
    $Chain.Build($Cert) | Out-Null
    $RootCert = $Chain.ChainElements | ForEach-Object { $_.Certificate } | Where-Object { $_.Subject -match "CN=Microsoft Root" }

    if ([string]::IsNullOrEmpty($RootCert)) {
        Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2
        return $false
    }

    $TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $_.Thumbprint -eq $RootCert.Thumbprint }
    if ([string]::IsNullOrEmpty($TrustedRoot)) {
        Write-LogEntry -Value "No trust found to root cert - aborting" -Severity 2
        return $false
    }

    Write-LogEntry -Value "Verified setup file signed by: $($Cert.Issuer)" -Severity 1
    return $true
}

function Test-ProofingToolsInstalled {
    <#
    .SYNOPSIS
        Returns $true when the proofing tools for the language are installed.
    .DESCRIPTION
        Only accepts proofing-specific evidence. The language appearing in InstalledLanguages
        or as a "<Product> - <lang>" entry means the LANGUAGE PACK is installed, which is a
        different Intune app. Accepting that was the false positive in version 2.0/2.1.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    # Source 1: the ProofingTools product's own culture subkey.
    foreach ($InstallID in Get-RegistrySubKeyNames64 -Key $C2RProductKey) {
        if (Test-RegistryKey64 -Key "$C2RProductKey\$InstallID\$ProofingProductID.16\$Language") {
            return $true
        }
    }

    # Source 2: the ".proof" ARP entry, e.g. "O365ProPlusRetail - nl-nl.proof".
    foreach ($ProductID in $BaseProductIDs) {
        if (Test-RegistryKey64 -Key "$UninstallKey\$ProductID - $Language.proof") {
            return $true
        }
    }

    # Source 3: any ".proof" ARP entry for this language, for product IDs not listed above.
    $ProofPattern = '\-\s*' + [regex]::Escape($Language) + '\.proof$'
    foreach ($KeyName in Get-RegistrySubKeyNames64 -Key $UninstallKey) {
        if ($KeyName -match $ProofPattern) {
            return $true
        }
    }

    return $false
}

function Wait-OfficeEngineIdle {
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

        Start-Sleep -Seconds $CheckIntervalSeconds
    }

    $Stopwatch.Stop()
    Write-LogEntry -Value "Timeout waiting for the Click-to-Run engine to become idle - continuing anyway" -Severity 2
    return $false
}

function Test-ProofingToolsStatus {
    <#
    .SYNOPSIS
        Waits until the proofing tools reaches the state expected for $Mode.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$LanguageID,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$Mode,

        [Parameter(Mandatory = $false)]
        [int]$MaxWaitMinutes = 30
    )

    Write-LogEntry -Value "Starting proofing tools $Mode verification for '$LanguageID' (max wait: $MaxWaitMinutes minutes)" -Severity 1

    $StartTime = Get-Date
    $TimeoutTime = $StartTime.AddMinutes($MaxWaitMinutes)
    $CheckIntervalSeconds = 15
    $StableCheckCount = 0
    $RequiredStableChecks = 2

    while ((Get-Date) -lt $TimeoutTime) {
        $ActiveProcesses = @(Get-Process -Name $OfficeSetupProcesses -ErrorAction SilentlyContinue)
        if ($ActiveProcesses.Count -gt 0) {
            Write-LogEntry -Value "Office setup still in progress (processes: $($ActiveProcesses.Name -join ', ')). Waiting..." -Severity 1
            $StableCheckCount = 0
            Start-Sleep -Seconds $CheckIntervalSeconds
            continue
        }

        $Installed = Test-ProofingToolsInstalled -Language $LanguageID
        $DesiredState = ($Mode -eq "Install")

        if ($Installed -eq $DesiredState) {
            $StableCheckCount++
            Write-LogEntry -Value "proofing tools '$LanguageID' is in the expected state for $Mode (stable check $StableCheckCount of $RequiredStableChecks)" -Severity 1

            if ($StableCheckCount -ge $RequiredStableChecks) {
                Write-LogEntry -Value "proofing tools '$LanguageID' $Mode verified successfully" -Severity 1
                return $true
            }
        }
        else {
            $StableCheckCount = 0
            Write-LogEntry -Value "proofing tools '$LanguageID' not yet in the expected state for $Mode. Waiting..." -Severity 1
        }

        $ElapsedMinutes = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 1)
        Write-LogEntry -Value "Elapsed time: $ElapsedMinutes minutes. Waiting $CheckIntervalSeconds seconds..." -Severity 1
        Start-Sleep -Seconds $CheckIntervalSeconds
    }

    Write-LogEntry -Value "Timeout reached after $MaxWaitMinutes minutes. proofing tools $Mode may not have completed." -Severity 3
    return $false
}

function Stop-ClickToRunService {
    try {
        $Service = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
        if ($Service -and $Service.Status -eq "Running") {
            Write-LogEntry -Value "Stopping Microsoft Office Click-to-Run service" -Severity 1
            Stop-Service -Name "ClickToRunSvc" -Force -ErrorAction Stop
            Write-LogEntry -Value "Microsoft Office Click-to-Run service stopped successfully" -Severity 1
        }
        else {
            Write-LogEntry -Value "Microsoft Office Click-to-Run service is not running or not found" -Severity 1
        }
    }
    catch {
        Write-LogEntry -Value "Failed to stop Microsoft Office Click-to-Run service: $($_.Exception.Message)" -Severity 2
    }
}

function Invoke-OfficeCleanup {
    if (-not (Test-Path -Path $SetupFolder)) { return }

    $ActiveProcesses = @(Get-Process -Name $OfficeSetupProcesses -ErrorAction SilentlyContinue)
    if ($ActiveProcesses.Count -gt 0) {
        Write-LogEntry -Value "Office setup processes still running ($($ActiveProcesses.Name -join ', ')) - skipping cleanup" -Severity 2
        return
    }

    try {
        Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Stop
        Write-LogEntry -Value "Setup folder cleaned up successfully" -Severity 1
    }
    catch {
        Write-LogEntry -Value "Warning: Could not remove setup folder: $($_.Exception.Message)" -Severity 2
    }
}
#endregion Functions

#region Main Script
$ExitCode = 0

try {
    # The Intune Management Extension is 32-bit, so without this the whole script runs under
    # WOW64 and every Office registry read is silently redirected to WOW6432Node.
    if ([System.Environment]::Is64BitOperatingSystem -and -not [System.Environment]::Is64BitProcess) {
        $NativePowerShell = Join-Path -Path $env:SystemRoot -ChildPath "SysNative\WindowsPowerShell\v1.0\powershell.exe"

        if (Test-Path -Path $NativePowerShell) {
            Write-LogEntry -Value "Running in a 32-bit host - relaunching in 64-bit PowerShell" -Severity 1

            $RelaunchArguments = @(
                "-NoProfile"
                "-NonInteractive"
                "-ExecutionPolicy", "Bypass"
                "-File", "`"$PSCommandPath`""
                "-LanguageID", $LanguageID
                "-Mode", $Mode
                "-MaxWaitMinutes", $MaxWaitMinutes
            )

            $Relaunched = Start-Process -FilePath $NativePowerShell -ArgumentList $RelaunchArguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
            Write-LogEntry -Value "64-bit instance completed with exit code: $($Relaunched.ExitCode)" -Severity 1
            exit $Relaunched.ExitCode
        }

        Write-LogEntry -Value "SysNative not available - continuing in the 32-bit host" -Severity 2
    }

    $FileName = if ($Mode -eq "Install") { "install.xml" } else { "uninstall.xml" }

    Write-LogEntry -Value "====" -Severity 1
    Write-LogEntry -Value "Initiating proofing tools '$LanguageID' $Mode process (v$ScriptVersion)" -Severity 1
    Write-LogEntry -Value "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name), 64-bit process: $([System.Environment]::Is64BitProcess)" -Severity 1
    Write-LogEntry -Value "Max wait time: $MaxWaitMinutes minutes" -Severity 1
    Write-LogEntry -Value "====" -Severity 1

    Invoke-OfficeCleanup
    $SetupFolder = (New-Item -ItemType Directory -Path "$env:SystemRoot\Temp" -Name OfficeSetup -Force).FullName

    # Download the evergreen Office Deployment Tool.
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"

    if (-not (Test-Path -Path $SetupFilePath)) {
        throw "Setup file not found after download"
    }
    Write-LogEntry -Value "Setup file found at $SetupFilePath" -Severity 1

    $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion
    Write-LogEntry -Value "Office C2R Setup is running version $OfficeCR2Version" -Severity 1

    if (-not (Invoke-FileCertVerification -FilePath $SetupFilePath)) {
        throw "Unable to verify setup file signature"
    }

    # Write the requested language into the local XML, then stage it next to setup.exe.
    $LocalXmlPath = Join-Path -Path $PSScriptRoot -ChildPath $FileName
    if (-not (Test-Path -Path $LocalXmlPath)) {
        throw "Local $FileName not found at: $LocalXmlPath"
    }

    Invoke-XMLUpdate -LanguageID $LanguageID -FileName $LocalXmlPath -Mode $Mode
    Copy-Item -Path $LocalXmlPath -Destination $SetupFolder -Force -ErrorAction Stop
    Write-LogEntry -Value "proofing tools '$LanguageID' configuration file copied" -Severity 1

    $StagedXmlPath = Join-Path -Path $SetupFolder -ChildPath $FileName

    # Never start setup.exe while the engine is mid-scenario, otherwise an uninstall issued
    # straight after an install silently does nothing until the device is rebooted.
    Wait-OfficeEngineIdle -TimeoutSeconds 900 -CheckIntervalSeconds 10 | Out-Null

    Write-LogEntry -Value "Starting proofing tools '$LanguageID' $Mode" -Severity 1
    $OfficeInstall = Start-Process -FilePath $SetupFilePath -ArgumentList "/configure `"$StagedXmlPath`"" -NoNewWindow -Wait -PassThru -ErrorAction Stop
    Write-LogEntry -Value "Setup.exe completed with exit code: $($OfficeInstall.ExitCode)" -Severity 1

    if ($OfficeInstall.ExitCode -ne 0) {
        Write-LogEntry -Value "proofing tools '$LanguageID' $Mode failed with exit code: $($OfficeInstall.ExitCode)" -Severity 3
        $ExitCode = 1
    }
    else {
        Write-LogEntry -Value "Setup.exe initiated successfully. Waiting for the $Mode to complete..." -Severity 1

        # This is the call that was broken in 2.1 (Test-ProofingToolsInstallation did not exist).
        $Verified = Test-ProofingToolsStatus -LanguageID $LanguageID -Mode $Mode -MaxWaitMinutes $MaxWaitMinutes

        if ($Verified) {
            Write-LogEntry -Value "proofing tools '$LanguageID' $Mode completed and verified successfully" -Severity 1
            if ($Mode -eq "Install") { Stop-ClickToRunService }
        }
        else {
            Write-LogEntry -Value "proofing tools '$LanguageID' $Mode verification failed or timed out" -Severity 3
            $ExitCode = 1
        }
    }
}
catch {
    Write-LogEntry -Value "Critical error during proofing tools $Mode : $($_.Exception.Message)" -Severity 3
    Write-LogEntry -Value "Stack trace: $($_.ScriptStackTrace)" -Severity 3
    $ExitCode = 1
}

Invoke-OfficeCleanup

Write-LogEntry -Value "====" -Severity 1
Write-LogEntry -Value "proofing tools '$LanguageID' $Mode completed with exit code: $ExitCode" -Severity 1
Write-LogEntry -Value "====" -Severity 1

exit $ExitCode
#endregion Main Script

