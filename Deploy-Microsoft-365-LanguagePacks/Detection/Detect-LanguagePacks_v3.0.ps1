<#
.SYNOPSIS
    Intune Win32 detection script for a Microsoft 365 Apps language pack.

.DESCRIPTION
    Confirms that a specific Office UI language is installed for an installed Click-to-Run
    product, and that the evidence is a real language pack rather than a proofing tools entry.

    Replaces v2.0, which had three problems:
      - Method 1 relied on the InstalledLanguages registry value. That value does not exist
        on current Click-to-Run builds, so the check silently never fired.
      - Method 3 looked for %ProgramFiles%\Microsoft Office\root\Office16\<lang>. Office uses
        numeric LCID folders there (1043, 1033), never "nl-nl", so that check never matched.
      - Diagnostics were written with Write-Output on both the success and failure paths.

    Detection rules (any one is sufficient, all are language-pack specific):
      1. ClickToRun\ProductReleaseIDs\<InstallID>\<Product>.16\<language> exists.
      2. An ARP entry named "<Product> - <language>" exists (and is not "<...>.proof").
      3. InstalledLanguages contains the language, on builds that still publish that value.

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

# The language pack this Intune app installs, in ODT format.
$LanguageID = 'nl-nl'

# Base Click-to-Run products the language pack can be attached to.
$BaseProductIDs = @('O365ProPlusRetail', 'O365BusinessRetail')

# Write a CMTrace log next to the Intune Management Extension logs.
$EnableLogFile = $true

#endregion ===================================================


#region Variables
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFileName = 'Detect-LanguagePacks.log'
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

function Get-InstalledOfficeProducts {
    $ProductReleaseIds = Get-RegistryValue64 -Key $C2RConfigKey -Name 'ProductReleaseIds'
    if ([string]::IsNullOrWhiteSpace($ProductReleaseIds)) { return @() }
    return @($ProductReleaseIds -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-LanguagePackInstalled {
    param(
        [Parameter(Mandatory = $true)][string]$Language
    )

    $InstalledProducts = Get-InstalledOfficeProducts
    if ($InstalledProducts.Count -eq 0) {
        Write-DetectionLog -Value 'No Click-to-Run product installed - a language pack cannot be present.' -Severity 2
        return $false
    }
    Write-DetectionLog -Value "Installed Click-to-Run products: [$($InstalledProducts -join ', ')]."

    # Only consider base products that are actually installed on this device.
    $Targets = @($BaseProductIDs | Where-Object { $InstalledProducts -contains $_ })
    if ($Targets.Count -eq 0) {
        Write-DetectionLog -Value "None of the configured base products [$($BaseProductIDs -join ', ')] are installed." -Severity 2
        return $false
    }

    foreach ($ProductID in $Targets) {
        # Source 1: per-install culture subkey.
        foreach ($InstallID in Get-RegistrySubKeyNames64 -Key $C2RProductKey) {
            if (Test-RegistryKey64 -Key "$C2RProductKey\$InstallID\$ProductID.16\$Language") {
                Write-DetectionLog -Value "Language [$Language] found under ProductReleaseIDs\$InstallID\$ProductID.16."
                return $true
            }
        }

        # Source 2: the per-language ARP entry. The ".proof" suffix is a proofing tools
        # entry, not a language pack, so it must not satisfy this check.
        $ArpKey = "$ProductID - $Language"
        if (Test-RegistryKey64 -Key "$UninstallKey\$ArpKey") {
            Write-DetectionLog -Value "Language [$Language] found as uninstall entry [$ArpKey]."
            return $true
        }
    }

    # Source 3: InstalledLanguages, present only on some builds.
    $InstalledLanguages = Get-RegistryValue64 -Key $C2RConfigKey -Name 'InstalledLanguages'
    if (-not [string]::IsNullOrWhiteSpace($InstalledLanguages)) {
        $Languages = @($InstalledLanguages -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($Languages -contains $Language) {
            Write-DetectionLog -Value "Language [$Language] found in InstalledLanguages [$InstalledLanguages]."
            return $true
        }
    }

    Write-DetectionLog -Value "Language pack [$Language] is not installed for any of [$($Targets -join ', ')]." -Severity 2
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
    Write-DetectionLog -Value '=== Microsoft 365 language pack detection started (v3.0) ==='
    Write-DetectionLog -Value "Target language [$LanguageID]."

    if (Test-LanguagePackInstalled -Language $LanguageID) {
        Write-DetectionLog -Value "=== DETECTED: language pack $LanguageID ==="
        Write-Output "Detected Microsoft 365 Apps language pack $LanguageID."
        exit 0
    }

    Write-DetectionLog -Value "=== NOT DETECTED: language pack $LanguageID ===" -Severity 2
    exit 1
}
catch {
    Write-DetectionLog -Value "=== DETECTION ERROR: $($_.Exception.Message) ===" -Severity 3
    Write-DetectionLog -Value "Stack trace: $($_.ScriptStackTrace)" -Severity 3
    exit 1
}
#endregion Main
