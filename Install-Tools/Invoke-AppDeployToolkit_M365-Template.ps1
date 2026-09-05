<#

.SYNOPSIS
PSAppDeployToolkit v4 wrapper for the Microsoft 365 Apps evergreen installer.

.DESCRIPTION
TEMPLATE. Copy this file into a PSADT package root as "Invoke-AppDeployToolkit.ps1".
See ..\Install-Tools\README.md for the full package build procedure.

It wraps Files\Install-Microsoft-365-Apps_v2.5.ps1 unchanged, so the same installer script,
Configuration.xml and Uninstall.xml serve both the plain Win32 script package and the PSADT
package. The wrapped script downloads the latest Office Deployment Tool setup.exe from the
evergreen URL on every run, so the package never goes stale.

Language handling stays entirely in Files\Configuration.xml (MatchOS + nl-nl + en-us);
nothing here changes it.

WHY THIS FILE EXISTS
The stock PSADT template is a skeleton: its Install/Uninstall functions contain empty
"<Perform Installation tasks here>" placeholders. This file fills those in with the
Microsoft 365 deployment logic, so every M365 package starts from the same tested wrapper
instead of being written by hand each time. It is a convenience, not a bug fix - the stock
template is not defective.

WHAT IS ADDED ON TOP OF THE STOCK TEMPLATE:
- Invoke-M365InstallerScript, Resolve-M365ExitCode, Show-M365CloseAppsPrompt and
  Write-M365InstalledVersion - the deployment logic itself.
- A -ProductID parameter, excluded from Open-ADTSession, which does not accept parameters
  it does not own.
- The wrapped script is launched through native-bitness PowerShell.
- The real ODT exit code is passed back to Intune instead of being masked as 60001.
- The module import path is "$PSScriptRoot\PSAppDeployToolkit\...", matching what
  New-ADTTemplate generates. If you build the package root with New-ADTTemplate you get
  this path already; it is spelled out here because copying the Frontend\v4 file by hand
  does not, and the result fails on a device with exit code 60008.

.PARAMETER DeploymentType
The type of deployment to perform.

.PARAMETER DeployMode
Interactive / Silent / NonInteractive / Auto. Leave on Auto for Intune: PSADT drops to
Silent automatically when no user is logged on or the device is in OOBE/Autopilot ESP.

.PARAMETER ProductID
Microsoft 365 Apps product to install/uninstall. Set the default below to match the package.

.PARAMETER SuppressRebootPassThru
Suppresses the 3010 return code from being passed back to the parent process.

.PARAMETER TerminalServerMode
Changes to "user install mode" for RDSH/Citrix servers.

.PARAMETER DisableLogging
Disables logging to file for the script.

.EXAMPLE
Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Auto

.EXAMPLE
Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent

.NOTES
Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Invoke-AppDeployToolkit.ps1/.exe
- 69000 - 69999: Recommended for user customized exit codes in Invoke-AppDeployToolkit.ps1
- 70000 - 79999: Recommended for user customized exit codes in the Extensions module.

Custom exit codes used by this script:
- 69001: The wrapped installer script was not found in the Files directory.

.LINK
https://psappdeploytoolkit.com

#>

[CmdletBinding()]
param
(
    # Default is 'Install'.
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [System.String]$DeploymentType,

    # Default is 'Auto'. Don't hard-code this unless required.
    [Parameter(Mandatory = $false)]
    [ValidateSet('Auto', 'Interactive', 'NonInteractive', 'Silent')]
    [System.String]$DeployMode,

    # Custom parameter - excluded from Open-ADTSession below.
    [Parameter(Mandatory = $false)]
    [ValidateSet('O365BusinessRetail', 'O365ProPlusRetail')]
    [System.String]$ProductID = 'O365BusinessRetail',

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$SuppressRebootPassThru,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$TerminalServerMode,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$DisableLogging
)


##================================================
## MARK: Variables
##================================================

# ---- EDIT THESE FOR EACH PACKAGE -------------------------------------------------------
# Name of the wrapped installer script inside the Files directory. Build-PSADTPackage.ps1
# fails validation if this file is not present in Files\.
$InstallerScriptName = 'Install-Microsoft-365-Apps_v2.5.ps1'

# How long the "please close these apps" countdown runs before the apps are closed for the
# user. Only ever shown when at least one Office app is actually running AND a user is
# logged on interactively. During Autopilot / ESP there is no user, so PSADT runs Silent
# and closes anything running without prompting.
$CloseAppsCountdownSeconds = 600
$CloseAppsCountdownSecondsUninstall = 60
# ----------------------------------------------------------------------------------------

$adtSession = @{
    # App variables.
    AppVendor = 'Microsoft'
    AppName = 'Microsoft 365 Apps for Business'
    AppVersion = 'Evergreen'
    AppArch = 'x64'
    AppLang = 'MUI'
    AppRevision = '01'
    AppSuccessExitCodes = @(0)
    AppRebootExitCodes = @(1641, 3010)
    AppProcessesToClose = @(
        @{ Name = 'winword'; Description = 'Microsoft Word' }
        @{ Name = 'excel'; Description = 'Microsoft Excel' }
        @{ Name = 'powerpnt'; Description = 'Microsoft PowerPoint' }
        @{ Name = 'outlook'; Description = 'Microsoft Outlook' }
        @{ Name = 'onenote'; Description = 'Microsoft OneNote' }
        @{ Name = 'onenotem'; Description = 'Microsoft OneNote' }
        @{ Name = 'msaccess'; Description = 'Microsoft Access' }
        @{ Name = 'mspub'; Description = 'Microsoft Publisher' }
        @{ Name = 'visio'; Description = 'Microsoft Visio' }
        @{ Name = 'winproj'; Description = 'Microsoft Project' }
        @{ Name = 'msoia'; Description = 'Office Telemetry Agent' }
    )
    AppScriptVersion = '1.0.0'
    AppScriptDate = '2026-09-04'
    AppScriptAuthor = 'UniFy-Endpoint'
    RequireAdmin = $true

    # Install Titles (Only set here to override defaults set by the toolkit).
    InstallName = ''
    InstallTitle = 'Microsoft 365 Apps'

    # Script variables.
    DeployAppScriptFriendlyName = $MyInvocation.MyCommand.Name
    DeployAppScriptParameters = $PSBoundParameters
    DeployAppScriptVersion = '4.1.8'
}


##================================================
## MARK: Helper
##================================================

function Get-M365PowerShellPath
{
    <#
    .SYNOPSIS
        Returns the path to the native-bitness Windows PowerShell 5.1 executable.

    .DESCRIPTION
        The wrapped installer script reads $env:ProgramFiles and HKLM:\SOFTWARE\Microsoft\Office.
        It has its own 64-bit relaunch guard, but launching it natively in the first place
        avoids the extra process hop and keeps the log readable.
    #>

    if ([System.Environment]::Is64BitOperatingSystem -and !([System.Environment]::Is64BitProcess))
    {
        return "$([System.Environment]::GetFolderPath('Windows'))\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    }
    return "$([System.Environment]::GetFolderPath('System'))\WindowsPowerShell\v1.0\powershell.exe"
}

function Invoke-M365InstallerScript
{
    <#
    .SYNOPSIS
        Runs the wrapped Install-Microsoft-365-Apps script and returns its exit code.
    #>

    [CmdletBinding()]
    [OutputType([System.Int32])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Install', 'Uninstall')]
        [System.String]$Mode
    )

    $scriptPath = Join-Path -Path $adtSession.DirFiles -ChildPath $InstallerScriptName
    if (!(Test-Path -LiteralPath $scriptPath -PathType Leaf))
    {
        Write-ADTLogEntry -Message "Wrapped installer script not found at [$scriptPath]." -Severity 3
        Close-ADTSession -ExitCode 69001
    }

    # Remove the mark-of-the-web in case the package was assembled from downloaded files.
    Unblock-File -LiteralPath $scriptPath -ErrorAction Ignore

    $psExe = Get-M365PowerShellPath
    $arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`" -Mode $Mode -ProductID $ProductID"

    Write-ADTLogEntry -Message "Launching [$psExe] $arguments"

    # -IgnoreExitCodes '*' hands the real ODT exit code back instead of letting a non-zero
    # code throw and be reported to Intune as the generic toolkit failure 60001.
    $result = Start-ADTProcess -FilePath $psExe -ArgumentList $arguments -WorkingDirectory $adtSession.DirFiles -CreateNoWindow -IgnoreExitCodes '*' -PassThru

    if (![System.String]::IsNullOrWhiteSpace($result.StdErr))
    {
        Write-ADTLogEntry -Message "Installer script stderr:`n$($result.StdErr)" -Severity 2
    }
    Write-ADTLogEntry -Message "Installer script returned exit code [$($result.ExitCode)]. Detailed log: %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log"
    return $result.ExitCode
}

function Resolve-M365ExitCode
{
    <#
    .SYNOPSIS
        Maps the wrapped script's exit code onto a PSADT/Intune-friendly session exit code.
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Int32]$ExitCode
    )

    if ($adtSession.AppSuccessExitCodes -contains $ExitCode)
    {
        return
    }
    if ($adtSession.AppRebootExitCodes -contains $ExitCode)
    {
        Write-ADTLogEntry -Message "A reboot is required to finish the $($adtSession.DeploymentType)."
        Close-ADTSession -ExitCode $ExitCode
    }
    Write-ADTLogEntry -Message "Microsoft 365 Apps $($adtSession.DeploymentType) failed with exit code [$ExitCode]." -Severity 3
    Close-ADTSession -ExitCode $ExitCode
}

function Show-M365CloseAppsPrompt
{
    <#
    .SYNOPSIS
        Shows the close-apps countdown, tolerating an empty AppProcessesToClose list.

    .DESCRIPTION
        Show-ADTInstallationWelcome rejects an empty -CloseProcesses array with a parameter
        validation error, which the outer handler turns into exit code 60001. Building the
        parameters as a splat means the app list can be trimmed to nothing without breaking
        the deployment.

        The dialog itself only appears when one of the listed apps is actually running and a
        user is logged on. Otherwise this returns immediately and the deployment continues.
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Int32]$CountdownSeconds,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CheckDiskSpace
    )

    $saiwParams = @{}
    if ($adtSession.AppProcessesToClose -and $adtSession.AppProcessesToClose.Count -gt 0)
    {
        $saiwParams.Add('CloseProcesses', $adtSession.AppProcessesToClose)
        $saiwParams.Add('CloseProcessesCountdown', $CountdownSeconds)
        $saiwParams.Add('PersistPrompt', $true)
    }
    if ($CheckDiskSpace)
    {
        $saiwParams.Add('CheckDiskSpace', $true)
        $saiwParams.Add('RequiredDiskSpace', 8192)
    }
    if ($saiwParams.Count -eq 0)
    {
        return
    }

    Show-ADTInstallationWelcome @saiwParams
}

function Write-M365InstalledVersion
{
    $version = Get-ADTRegistryKey -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'VersionToReport'
    $products = Get-ADTRegistryKey -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'ProductReleaseIds'
    $languages = Get-ADTRegistryKey -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'ClientCulture'

    if ($version)
    {
        Write-ADTLogEntry -Message "Office Click-to-Run version [$version], products [$products], culture [$languages]."
    }
    else
    {
        Write-ADTLogEntry -Message 'Office Click-to-Run configuration key not present.' -Severity 2
    }
}


##================================================
## MARK: Deployment
##================================================

function Install-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Close any running Office apps.
    ## - No deferral. Intune treats PSADT's deferral exit code (60012) as a failure unless it
    ##   is mapped to "Retry", and a deferred install during Autopilot would strand the device.
    ## - If no Office app is running, this returns immediately and the install proceeds.
    ## - If an Office app IS running and a user is logged on, the user gets a countdown dialog
    ##   listing the apps, with a "Close Programs" button; at zero the apps are closed for them.
    ## - With no logged-on user (Autopilot / ESP / device reset) PSADT runs Silent and closes
    ##   anything running without prompting.
    Show-M365CloseAppsPrompt -CountdownSeconds $CloseAppsCountdownSeconds -CheckDiskSpace

    ## Show Progress Message.
    Show-ADTInstallationProgress -StatusMessage 'Installing Microsoft 365 Apps. This can take 15-45 minutes depending on your connection. Please wait...'


    ##================================================
    ## MARK: Install
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    Write-ADTLogEntry -Message "Installing Microsoft 365 Apps product [$ProductID] using the evergreen Office Deployment Tool."
    $exitCode = Invoke-M365InstallerScript -Mode Install


    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    Resolve-M365ExitCode -ExitCode $exitCode
    Write-M365InstalledVersion
}

function Uninstall-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    Show-M365CloseAppsPrompt -CountdownSeconds $CloseAppsCountdownSecondsUninstall

    ## Show Progress Message.
    Show-ADTInstallationProgress -StatusMessage 'Removing Microsoft 365 Apps. Please wait...'


    ##================================================
    ## MARK: Uninstall
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    Write-ADTLogEntry -Message "Uninstalling Microsoft 365 Apps product [$ProductID]."
    $exitCode = Invoke-M365InstallerScript -Mode Uninstall


    ##================================================
    ## MARK: Post-Uninstallation
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    Resolve-M365ExitCode -ExitCode $exitCode
}

function Repair-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Repair
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    Show-M365CloseAppsPrompt -CountdownSeconds $CloseAppsCountdownSecondsUninstall

    ## Show Progress Message.
    Show-ADTInstallationProgress -StatusMessage 'Repairing Microsoft 365 Apps. Please wait...'


    ##================================================
    ## MARK: Repair
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## Re-running /configure against the same configuration.xml repairs/reconciles the install.
    Write-ADTLogEntry -Message "Repairing Microsoft 365 Apps product [$ProductID] by re-applying configuration.xml."
    $exitCode = Invoke-M365InstallerScript -Mode Install


    ##================================================
    ## MARK: Post-Repair
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    Resolve-M365ExitCode -ExitCode $exitCode
    Write-M365InstalledVersion
}


##================================================
## MARK: Initialization
##================================================

# Set strict error handling across entire operation.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

# Import the module and instantiate a new session.
try
{
    # NOTE: the stock Frontend\v4 template uses "$PSScriptRoot\..\..\..\PSAppDeployToolkit".
    # In a package root the module sits directly beside this file, so the path is shortened.
    if (Test-Path -LiteralPath "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -PathType Leaf)
    {
        Get-ChildItem -LiteralPath "$PSScriptRoot\PSAppDeployToolkit" -Recurse -File | Unblock-File -ErrorAction Ignore
        Import-Module -FullyQualifiedName @{ ModuleName = "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
    }
    else
    {
        Import-Module -FullyQualifiedName @{ ModuleName = 'PSAppDeployToolkit'; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
    }

    # Open a new deployment session, replacing $adtSession with a DeploymentSession.
    # ProductID is ours, not the toolkit's, so it must be excluded from the session parameters.
    $iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation -Exclude ProductID
    $adtSession = Remove-ADTHashtableNullOrEmptyValues -Hashtable $adtSession
    $adtSession = Open-ADTSession @adtSession @iadtParams -PassThru
}
catch
{
    $Host.UI.WriteErrorLine((Out-String -InputObject $_ -Width ([System.Int32]::MaxValue)))
    exit 60008
}


##================================================
## MARK: Invocation
##================================================

# Commence the actual deployment operation.
try
{
    # Import any found extensions before proceeding with the deployment.
    Get-ChildItem -LiteralPath $PSScriptRoot -Directory | & {
        process
        {
            if ($_.Name -match 'PSAppDeployToolkit\..+$')
            {
                Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
                Import-Module -Name $_.FullName -Force
            }
        }
    }

    # Invoke the deployment and close out the session.
    & "$($adtSession.DeploymentType)-ADTDeployment"
    Close-ADTSession
}
catch
{
    # An unhandled error has been caught.
    $mainErrorMessage = "An unhandled error within [$($MyInvocation.MyCommand.Name)] has occurred.`n$(Resolve-ADTErrorRecord -ErrorRecord $_)"
    Write-ADTLogEntry -Message $mainErrorMessage -Severity 3
    Close-ADTSession -ExitCode 60001
}

