<#
.SYNOPSIS
    Intune Win32 detection script for Microsoft 365 Apps (Click-to-Run).

.DESCRIPTION
    Confirms that ONE specific Microsoft 365 Apps product is installed, that it is at or
    above an optional minimum version, and that every required language is present.

    Replaces v1.0, which used a "2 out of 4 weak checks" heuristic. That heuristic reported
    O365BusinessRetail as detected on a machine that only had O365ProPlusRetail, because it
    matched either product ID and accepted "any DisplayName like *Microsoft Office*" and
    "the ClickToRunSvc service exists" as evidence.

    Detection rules (ALL must pass):
      1. Office Click-to-Run is installed (VersionToReport present + OfficeClickToRun.exe on disk).
      2. ProductReleaseIds contains the EXACT target product ID (comma-split, exact match).
      3. Every language in $RequiredLanguages is installed for that product.
      4. VersionToReport is >= $MinimumVersion, when a minimum is configured.

    All registry reads use the 64-bit view explicitly, so the result is identical whether
    Intune runs this in a 32-bit or 64-bit host.

    Intune treats an app as detected only when the script exits 0 AND writes to STDOUT.
    Therefore exactly one line is written to STDOUT on success, and nothing on failure.
    Diagnostics go to the log file and to the verbose stream.

.NOTES
    Version: 2.0
    Author:  UniFy-Endpoint

    Intune configuration:
      Detection method                            : Use a custom detection script
      Run script as 32-bit process on 64-bit client: No
      Enforce script signature check               : No
#>

[CmdletBinding()]
param()

#region ===================== CONFIGURE ME =====================

# The exact Click-to-Run product this Intune app installs.
# O365BusinessRetail = Microsoft 365 Apps for business
# O365ProPlusRetail  = Microsoft 365 Apps for enterprise
$TargetProductID = 'O365BusinessRetail'

# Languages that Configuration.xml always installs. Detection fails if any is missing,
# which makes Intune re-run the install and repair the language state.
# Set to @() to skip the language check entirely.
$RequiredLanguages = @('nl-nl', 'en-us')

# Optional minimum build, e.g. '16.0.17928.20114'. Leave empty to accept any version.
$MinimumVersion = ''

# Write a CMTrace log next to the Intune Management Extension logs.
$EnableLogFile = $true

#endregion ===================================================


#region Variables
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFileName = 'Detect-Microsoft-365-Apps.log'
$LogFilePath = Join-Path -Path $LogFolder -ChildPath $LogFileName
$LogMaxBytes = 1MB

$C2RConfigKey = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$C2RProductKey = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
$UninstallKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
#endregion Variables


#region Functions
function Write-DetectionLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('1', '2', '3')]
        [string]$Severity = '1'
    )

    Write-Verbose -Message $Value

    if (-not $EnableLogFile) { return }

    try {
        if (-not (Test-Path -LiteralPath $LogFolder)) {
            New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        }

        # Keep the log from growing without bound - detection runs every few hours.
        $ExistingLog = Get-Item -LiteralPath $LogFilePath -ErrorAction SilentlyContinue
        if ($ExistingLog -and $ExistingLog.Length -gt $LogMaxBytes) {
            Remove-Item -LiteralPath $LogFilePath -Force -ErrorAction SilentlyContinue
        }

        $Time = '{0} {1}' -f (Get-Date -Format 'HH:mm:ss.fff'), $script:TimezoneBias
        $Date = Get-Date -Format 'MM-dd-yyyy'
        $Context = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $LogText = "<![LOG[$Value]LOG]!><time=""$Time"" date=""$Date"" component=""$LogFileName"" context=""$Context"" type=""$Severity"" thread=""$PID"" file="""">"

        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch {
        # Logging must never influence the detection result.
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

function Test-OfficeClickToRunPresent {
    $Version = Get-RegistryValue64 -Key $C2RConfigKey -Name 'VersionToReport'
    if ([string]::IsNullOrWhiteSpace($Version)) {
        Write-DetectionLog -Value 'Click-to-Run VersionToReport is missing - Office is not installed.' -Severity 2
        return $false
    }

    # $env:ProgramFiles points at Program Files (x86) inside a 32-bit host, so prefer ProgramW6432.
    $ProgramFiles64 = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $C2RExe = Join-Path -Path $ProgramFiles64 -ChildPath 'Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe'

    if (-not (Test-Path -LiteralPath $C2RExe)) {
        Write-DetectionLog -Value "Registry reports version [$Version] but [$C2RExe] is missing." -Severity 2
        return $false
    }

    Write-DetectionLog -Value "Office Click-to-Run present, VersionToReport [$Version]."
    return $true
}

function Test-OfficeProductInstalled {
    param(
        [Parameter(Mandatory = $true)][string]$ProductID
    )

    $ProductReleaseIds = Get-RegistryValue64 -Key $C2RConfigKey -Name 'ProductReleaseIds'
    if ([string]::IsNullOrWhiteSpace($ProductReleaseIds)) {
        Write-DetectionLog -Value 'ProductReleaseIds is empty - no Click-to-Run product registered.' -Severity 2
        return $false
    }

    Write-DetectionLog -Value "ProductReleaseIds reports [$ProductReleaseIds]."
    $Installed = @($ProductReleaseIds -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    if ($Installed -contains $ProductID) {
        Write-DetectionLog -Value "Target product [$ProductID] is installed."
        return $true
    }

    Write-DetectionLog -Value "Target product [$ProductID] is NOT installed. Present: [$($Installed -join ', ')]." -Severity 2
    return $false
}

function Test-OfficeLanguageInstalled {
    param(
        [Parameter(Mandatory = $true)][string]$ProductID,
        [Parameter(Mandatory = $true)][string]$Language
    )

    # Source 1: the per-install culture subkeys, e.g.
    # ClickToRun\ProductReleaseIDs\<InstallID>\O365ProPlusRetail.16\nl-nl
    foreach ($InstallID in Get-RegistrySubKeyNames64 -Key $C2RProductKey) {
        if (Test-RegistryKey64 -Key "$C2RProductKey\$InstallID\$ProductID.16\$Language") {
            Write-DetectionLog -Value "Language [$Language] found under ProductReleaseIDs\$InstallID\$ProductID.16."
            return $true
        }
    }

    # Source 2: the per-language ARP entry, e.g. "O365ProPlusRetail - nl-nl".
    if (Test-RegistryKey64 -Key "$UninstallKey\$ProductID - $Language") {
        Write-DetectionLog -Value "Language [$Language] found as uninstall entry [$ProductID - $Language]."
        return $true
    }

    # Source 3: InstalledLanguages, which only exists on some builds.
    $InstalledLanguages = Get-RegistryValue64 -Key $C2RConfigKey -Name 'InstalledLanguages'
    if (-not [string]::IsNullOrWhiteSpace($InstalledLanguages)) {
        $Languages = @($InstalledLanguages -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($Languages -contains $Language) {
            Write-DetectionLog -Value "Language [$Language] found in InstalledLanguages [$InstalledLanguages]."
            return $true
        }
    }

    Write-DetectionLog -Value "Language [$Language] is NOT installed for [$ProductID]." -Severity 2
    return $false
}

function Test-OfficeVersion {
    param(
        [Parameter(Mandatory = $true)][string]$Minimum
    )

    $Current = Get-RegistryValue64 -Key $C2RConfigKey -Name 'VersionToReport'

    try {
        $CurrentVersion = [version]$Current
        $MinimumVersionObject = [version]$Minimum
    }
    catch {
        Write-DetectionLog -Value "Could not compare version [$Current] against minimum [$Minimum] - treating as not detected." -Severity 3
        return $false
    }

    if ($CurrentVersion -ge $MinimumVersionObject) {
        Write-DetectionLog -Value "Version [$Current] meets the minimum [$Minimum]."
        return $true
    }

    Write-DetectionLog -Value "Version [$Current] is below the minimum [$Minimum]." -Severity 2
    return $false
}
#endregion Functions


#region Main
try {
    $script:TimezoneBias = (Get-CimInstance -ClassName Win32_TimeZone -ErrorAction SilentlyContinue).Bias
}
catch {
    $script:TimezoneBias = 0
}

try {
    Write-DetectionLog -Value '=== Microsoft 365 Apps detection started (v2.0) ==='
    Write-DetectionLog -Value "Target product [$TargetProductID], required languages [$($RequiredLanguages -join ', ')], minimum version [$MinimumVersion]."

    $Detected = $true

    if (-not (Test-OfficeClickToRunPresent)) {
        $Detected = $false
    }

    if ($Detected -and -not (Test-OfficeProductInstalled -ProductID $TargetProductID)) {
        $Detected = $false
    }

    if ($Detected -and $RequiredLanguages.Count -gt 0) {
        foreach ($Language in $RequiredLanguages) {
            if (-not (Test-OfficeLanguageInstalled -ProductID $TargetProductID -Language $Language)) {
                $Detected = $false
            }
        }
    }

    if ($Detected -and -not [string]::IsNullOrWhiteSpace($MinimumVersion)) {
        if (-not (Test-OfficeVersion -Minimum $MinimumVersion)) {
            $Detected = $false
        }
    }

    # Informational only - never gates the result, otherwise a background Office update
    # would make a healthy install disappear from Intune.
    $ExecutingScenario = Get-RegistryValue64 -Key $C2RConfigKey -Name 'ExecutingScenario'
    if (-not [string]::IsNullOrWhiteSpace($ExecutingScenario)) {
        Write-DetectionLog -Value "Note: a Click-to-Run operation is currently running [$ExecutingScenario]." -Severity 2
    }

    if ($Detected) {
        $Version = Get-RegistryValue64 -Key $C2RConfigKey -Name 'VersionToReport'
        Write-DetectionLog -Value "=== DETECTED: $TargetProductID $Version ==="
        # The single STDOUT line Intune needs to mark the app as installed.
        Write-Output "Detected $TargetProductID version $Version with languages $($RequiredLanguages -join ', ')."
        exit 0
    }

    Write-DetectionLog -Value "=== NOT DETECTED: $TargetProductID ===" -Severity 2
    exit 1
}
catch {
    Write-DetectionLog -Value "=== DETECTION ERROR: $($_.Exception.Message) ===" -Severity 3
    Write-DetectionLog -Value "Stack trace: $($_.ScriptStackTrace)" -Severity 3
    exit 1
}
#endregion Main
