<#
.SYNOPSIS
    Intune Win32 detection script for Microsoft 365 Apps proofing tools.

.DESCRIPTION
    Confirms that the proofing tools for a specific language are installed. Proofing tools
    are a separate Click-to-Run product from the language pack, so this script only accepts
    proofing-specific evidence.

    Replaces v2.0, whose Method 1 returned "installed" as soon as the language appeared in
    InstalledLanguages. A full language pack also adds the language there, so the proofing
    tools app could report as installed when only the language pack was present.
    v2.0 also probed Office16\Proof\<lang>, which is not where Click-to-Run stores proofing
    files, and wrote diagnostics to STDOUT on both the success and failure paths.

    Detection rules (any one is sufficient):
      1. ClickToRun\ProductReleaseIDs\<InstallID>\ProofingTools.16\<language> exists.
      2. An ARP entry named "<Product> - <language>.proof" exists.
      3. ProductReleaseIds contains ProofingTools AND the language culture key exists for it.

    All registry reads use the 64-bit view explicitly.

    Intune treats an app as detected only when the script exits 0 AND writes to STDOUT.
    Exactly one line goes to STDOUT on success, nothing on failure.

.NOTES
    Version: 3.0
    Author:  UniFy-Endpoint

    Intune configuration:
      Detection method                            : Use a custom detection script
      Run script as 32-bit process on 64-bit client: No
      Enforce script signature check               : No
#>

[CmdletBinding()]
param()

#region ===================== CONFIGURE ME =====================

# The proofing tools language this Intune app installs, in ODT format.
$LanguageID = 'nl-nl'

# Base Click-to-Run products the proofing tools can be attached to.
$BaseProductIDs = @('O365ProPlusRetail', 'O365BusinessRetail')

# Write a CMTrace log next to the Intune Management Extension logs.
$EnableLogFile = $true

#endregion ===================================================


#region Variables
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFileName = 'Detect-ProofingTools.log'
$LogFilePath = Join-Path -Path $LogFolder -ChildPath $LogFileName
$LogMaxBytes = 1MB

$C2RConfigKey = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$C2RProductKey = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
$UninstallKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
$ProofingProductID = 'ProofingTools'
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

function Test-ProofingToolsInstalled {
    param(
        [Parameter(Mandatory = $true)][string]$Language
    )

    # Source 1: the ProofingTools product's own culture subkey.
    foreach ($InstallID in Get-RegistrySubKeyNames64 -Key $C2RProductKey) {
        if (Test-RegistryKey64 -Key "$C2RProductKey\$InstallID\$ProofingProductID.16\$Language") {
            Write-DetectionLog -Value "Proofing tools [$Language] found under ProductReleaseIDs\$InstallID\$ProofingProductID.16."
            return $true
        }
    }

    # Source 2: the ".proof" ARP entry, e.g. "O365ProPlusRetail - nl-nl.proof".
    foreach ($ProductID in $BaseProductIDs) {
        $ArpKey = "$ProductID - $Language.proof"
        if (Test-RegistryKey64 -Key "$UninstallKey\$ArpKey") {
            Write-DetectionLog -Value "Proofing tools found as uninstall entry [$ArpKey]."
            return $true
        }
    }

    # Source 3: any ".proof" ARP entry for this language, for product IDs not listed above.
    $ProofPattern = '\-\s*' + [regex]::Escape($Language) + '\.proof$'
    foreach ($KeyName in Get-RegistrySubKeyNames64 -Key $UninstallKey) {
        if ($KeyName -match $ProofPattern) {
            Write-DetectionLog -Value "Proofing tools found as uninstall entry [$KeyName]."
            return $true
        }
    }

    $ProductReleaseIds = Get-RegistryValue64 -Key $C2RConfigKey -Name 'ProductReleaseIds'
    Write-DetectionLog -Value "Proofing tools for [$Language] not installed. ProductReleaseIds: [$ProductReleaseIds]." -Severity 2
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
    Write-DetectionLog -Value '=== Microsoft 365 proofing tools detection started (v3.0) ==='
    Write-DetectionLog -Value "Target language [$LanguageID]."

    if (Test-ProofingToolsInstalled -Language $LanguageID) {
        Write-DetectionLog -Value "=== DETECTED: proofing tools $LanguageID ==="
        Write-Output "Detected Microsoft 365 Apps proofing tools $LanguageID."
        exit 0
    }

    Write-DetectionLog -Value "=== NOT DETECTED: proofing tools $LanguageID ===" -Severity 2
    exit 1
}
catch {
    Write-DetectionLog -Value "=== DETECTION ERROR: $($_.Exception.Message) ===" -Severity 3
    Write-DetectionLog -Value "Stack trace: $($_.ScriptStackTrace)" -Severity 3
    exit 1
}
#endregion Main
