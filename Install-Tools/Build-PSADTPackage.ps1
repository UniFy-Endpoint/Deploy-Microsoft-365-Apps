<#
.SYNOPSIS
    Validates PSAppDeployToolkit (PSADT v4) package folders before packaging, and optionally
    builds the .intunewin with the Microsoft Win32 Content Prep Tool.

.DESCRIPTION
    WHAT THIS IS
    ------------
    A read-only pre-flight check for PSADT v4 packages destined for Intune. It catches the
    assembly mistakes that produce a package which looks perfectly fine on the packaging
    workstation and only fails once it reaches a real device - by which point you are reading
    IME logs instead of packaging the next app.

    By default it COPIES NOTHING and CHANGES NOTHING - it only reports. The one exception is
    -Restore, which writes the toolkit files into each package root; that is stated plainly in
    the output when it happens.

    It is a workstation tool and must never be included inside a package you deploy.

    FIRST RUN AFTER CLONING
    -----------------------
    The PSAppDeployToolkit module is about 19 MB per package and is deliberately NOT committed
    to this repository. A fresh clone therefore has the wrapper, the payload and the detection
    script, but no toolkit. Run this once to fetch and stage it:

        .\Build-PSADTPackage.ps1 -Restore

    That downloads PSAppDeployToolkit from the PowerShell Gallery into a .tools cache beside
    the app folders, copies it into each package root, and validates the result. Add
    -BuildIntuneWin to also download IntuneWinAppUtil.exe and produce the .intunewin files.

    Downloads happen only when a tool is not already cached. -NoDownload forbids them
    entirely, for an air-gapped workstation.

    EXPECTED FOLDER LAYOUT
    ----------------------
    This script assumes each app folder contains an AppInstaller subfolder that IS the PSADT
    package root - that is, the folder you hand to IntuneWinAppUtil:

        <SolutionRoot>\
          <AppName>\
            AppInstaller\                    <- PSADT package root
              Invoke-AppDeployToolkit.ps1      the wrapper you customised
              Invoke-AppDeployToolkit.exe      the Intune "setup file"
              PSAppDeployToolkit\              the toolkit module
              PSAppDeployToolkit.Extensions\
              Files\                           your installer script + config files
              SupportFiles\
            Detection\                       <- detection script, uploaded separately
            Package\                         <- .intunewin output

    An app folder whose AppInstaller has no Invoke-AppDeployToolkit.ps1 is treated as a plain
    Win32 script package and reported as "skipped". That is a normal state, not a failure.

    If your packages use a different folder name for the package root, change the two
    'AppInstaller' literals in Find-SolutionRoot and Test-PSADTPackage.

    WHAT IT CHECKS, AND WHY EACH ONE MATTERS
    ----------------------------------------
      1. Invoke-AppDeployToolkit.exe is present.
         It is the setup file you name in Intune. Without it there is nothing to launch.

      2. Invoke-AppDeployToolkit.ps1 is present and has no PowerShell syntax errors.
         A syntax error surfaces on the device as toolkit exit code 60008, with no useful
         message in the Intune portal.

      3. The toolkit module is present at AppInstaller\PSAppDeployToolkit\.
         Intune extracts the package to a device that has no PSADT installed, so the module
         has to travel inside the package.

      4. The wrapper's module import points at "$PSScriptRoot\PSAppDeployToolkit\...".
         THIS IS THE MOST COMMON PACKAGING MISTAKE, and it is not a defect in PSADT. The
         template in the toolkit's Frontend\v4 folder imports from
         "$PSScriptRoot\..\..\..\PSAppDeployToolkit", which resolves correctly from where that
         file normally sits. PSADT's own New-ADTTemplate cmdlet rewrites the path when it
         generates a package root. Copy the file by hand instead and you skip that rewrite:
         the import silently falls through to PSModulePath, and the deployment dies with exit
         code 60008 on every device that does not already have PSADT installed.

      5. The script named by $InstallerScriptName in the wrapper actually exists in Files\.
         Catches bumping your installer to a new version number and forgetting to update the
         wrapper. On a device this shows up as a custom exit code or a "file not found".

      6. Files\ is not empty and has no nested package folder inside it.
         An empty Files\ means the payload was never copied in.

      7. No path inside the package exceeds 240 characters.
         PSADT's Strings and lib trees are deep. If your working folder is already long, the
         toolkit's own DLLs go past MAX_PATH and fail to load - again as exit code 60008.
         Intune extracts to C:\Windows\IMECache\, so this only bites on the workstation.

    PREREQUISITES
    -------------
      - Windows PowerShell 5.1 or PowerShell 7+.
      - No administrator rights required.
      - Internet access on first run, to fetch PSAppDeployToolkit and IntuneWinAppUtil.exe.
        Both are cached under .tools and reused afterwards. Use -NoDownload plus
        -IntuneWinAppUtilPath to work fully offline.
      - No PowerShellGet or NuGet provider needed; the toolkit is fetched as a plain zip.

    HOW TO READ THE OUTPUT
    ----------------------
      [OK  ]  check passed
      [WARN]  informational; does not fail the run (e.g. a Win32-only app was skipped)
      [FAIL]  the package would not work as-is; each FAIL line is followed by a "Fix:" line

    Exit code 0 means every package checked is good. Exit code 1 means at least one failed.
    Suitable for use as a CI gate.

.PARAMETER SolutionRoot
    The folder that CONTAINS the app folders. If omitted, the script walks up from its own
    location until it finds a folder holding at least one app folder, so it works whether you
    keep it beside the app folders or in a tools subfolder.

.PARAMETER PackagePath
    One or more specific app folders to check. Accepts relative paths. If omitted, every app
    folder under the solution root is checked.

.PARAMETER Restore
    Download PSAppDeployToolkit if it is not already cached, then stage it into each package
    root. Run this once after cloning the repository. Existing toolkit files are left alone
    unless -Force is also given. This is the only switch that writes into a package.

.PARAMETER Force
    With -Restore, replace toolkit files that are already present instead of leaving them.
    Use it to move a package onto a different PSADT version.

.PARAMETER ToolsPath
    Where downloaded tools are cached. Defaults to a .tools folder beside the app folders.
    Add it to .gitignore - it holds third-party binaries that do not belong in the repository.

.PARAMETER PSADTVersion
    Which PSAppDeployToolkit version to download. Pinned by default so a clone of this
    repository builds the same package today and in a year. Check
    https://www.powershellgallery.com/packages/PSAppDeployToolkit for newer releases.

.PARAMETER IntuneWinAppUtilVersion
    Which tag of Microsoft-Win32-Content-Prep-Tool to download IntuneWinAppUtil.exe from.

.PARAMETER NoDownload
    Never reach the internet. Fails with a clear message if a needed tool is not already
    cached. Combine with -IntuneWinAppUtilPath on an air-gapped workstation.

.PARAMETER BuildIntuneWin
    After a package passes its checks, run IntuneWinAppUtil.exe on it and drop the .intunewin
    in that app's Package\ folder. A package that fails its checks is never built.

.PARAMETER IntuneWinAppUtilPath
    Full path to IntuneWinAppUtil.exe. Only needed with -BuildIntuneWin, and only if the tool
    is not already on your PATH.

.EXAMPLE
    .\Build-PSADTPackage.ps1 -Restore

    FIRST RUN AFTER CLONING. Downloads PSAppDeployToolkit, stages it into every package root,
    and validates the result.

.EXAMPLE
    .\Build-PSADTPackage.ps1 -Restore -BuildIntuneWin

    The whole thing from a fresh clone: fetch both tools, stage the packages, validate, and
    produce the .intunewin files. No paths to supply.

.EXAMPLE
    .\Build-PSADTPackage.ps1

    Check every app folder found. Nothing is written or built.

.EXAMPLE
    .\Build-PSADTPackage.ps1 -PackagePath ..\Microsoft-365-Apps-Enterprise

    Check a single app.

.EXAMPLE
    .\Build-PSADTPackage.ps1 -PackagePath ..\Microsoft-365-Apps-Business `
        -BuildIntuneWin -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe

    Check, then build Microsoft-365-Apps-Business\Package\Invoke-AppDeployToolkit.intunewin.

.EXAMPLE
    .\Build-PSADTPackage.ps1 -SolutionRoot D:\Packaging\MyApps -BuildIntuneWin

    Point at a different set of app folders, with IntuneWinAppUtil.exe already on PATH.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. Progress is written to the host; the result is conveyed by the exit code.

.NOTES
    Version: 3.0
    Author:  UniFy-Endpoint
    Updated: 04-09-2026

.LINK
    https://psappdeploytoolkit.com

.LINK
    https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SolutionRoot,

    [Parameter(Mandatory = $false)]
    [string[]]$PackagePath,

    [Parameter(Mandatory = $false)]
    [switch]$Restore,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$BuildIntuneWin,

    [Parameter(Mandatory = $false)]
    [string]$IntuneWinAppUtilPath,

    [Parameter(Mandatory = $false)]
    [string]$ToolsPath,

    [Parameter(Mandatory = $false)]
    [string]$PSADTVersion = '4.1.8',

    [Parameter(Mandatory = $false)]
    [string]$IntuneWinAppUtilVersion = 'v1.8.7',

    [Parameter(Mandatory = $false)]
    [switch]$NoDownload
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message, [string]$Status = 'INFO')

    $Colour = switch ($Status) {
        'OK' { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host ('[{0,-4}] {1}' -f $Status, $Message) -ForegroundColor $Colour
}

function Find-SolutionRoot {
    <#
    .SYNOPSIS
        Finds the folder that holds the app folders.
    .DESCRIPTION
        This script lives in Install-Tools\, a sibling of the app folders rather than their
        parent, so $PSScriptRoot is not the solution root. Walk up until a folder is found
        that directly contains at least one app folder.
    #>
    param([string]$StartPath)

    $Current = Get-Item -LiteralPath $StartPath
    for ($Depth = 0; $Depth -lt 5 -and $Current; $Depth++) {
        $Found = @(Get-ChildItem -LiteralPath $Current.FullName -Directory -ErrorAction SilentlyContinue |
                Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'AppInstaller') })

        if ($Found.Count -gt 0) {
            return $Current.FullName
        }
        $Current = $Current.Parent
    }

    return $null
}

function Assert-MicrosoftSignature {
    <#
    .SYNOPSIS
        Refuses to use a downloaded executable that is not validly signed by Microsoft.
    #>
    param([string]$Path, [string]$ExpectedSubject = 'O=Microsoft Corporation')

    $Signature = Get-AuthenticodeSignature -FilePath $Path
    if ($Signature.Status -ne 'Valid' -or $Signature.SignerCertificate.Subject -notmatch [regex]::Escape($ExpectedSubject)) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        throw "Downloaded file failed signature validation and was deleted. Status: $($Signature.Status); Signer: $($Signature.SignerCertificate.Subject)"
    }
    Write-Step "signature verified: $($Signature.SignerCertificate.Subject.Split(',')[0])" 'OK'
}

function Get-PSADTToolkit {
    <#
    .SYNOPSIS
        Returns the path to a local copy of the PSAppDeployToolkit module, downloading it
        from the PowerShell Gallery on first use.
    .DESCRIPTION
        The module is NOT committed to this repository, so a fresh clone has to fetch it.
        It is cached under $ToolsPath and reused on later runs.

        The download is a plain HTTPS GET of the Gallery's .nupkg, which is an ordinary zip.
        That avoids depending on PowerShellGet or the NuGet provider being present and
        configured, which is not a safe assumption on a locked-down packaging workstation.
    #>
    param([string]$ToolsPath, [string]$Version, [switch]$NoDownload)

    $Target = Join-Path -Path $ToolsPath -ChildPath "PSAppDeployToolkit-$Version"
    $Manifest = Join-Path -Path $Target -ChildPath 'PSAppDeployToolkit.psd1'

    if (Test-Path -LiteralPath $Manifest) {
        Write-Step "PSAppDeployToolkit $Version already cached at $Target" 'OK'
        return $Target
    }

    if ($NoDownload) {
        throw "PSAppDeployToolkit $Version is not cached at $Target and -NoDownload was specified."
    }

    Write-Step "Downloading PSAppDeployToolkit $Version from the PowerShell Gallery..."
    New-Item -Path $ToolsPath -ItemType Directory -Force | Out-Null
    $Nupkg = Join-Path -Path $ToolsPath -ChildPath "psadt-$Version.zip"
    $Uri = "https://www.powershellgallery.com/api/v2/package/PSAppDeployToolkit/$Version"

    $OldProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Uri -OutFile $Nupkg -UseBasicParsing -TimeoutSec 300
    }
    finally {
        $ProgressPreference = $OldProgress
    }

    if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Recurse -Force }
    Expand-Archive -LiteralPath $Nupkg -DestinationPath $Target -Force
    Remove-Item -LiteralPath $Nupkg -Force -ErrorAction SilentlyContinue

    # Strip the NuGet packaging artefacts so what is left is just the module.
    foreach ($Artefact in '_rels', 'package') {
        $p = Join-Path -Path $Target -ChildPath $Artefact
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    }
    Get-ChildItem -LiteralPath $Target -File |
        Where-Object { $_.Name -eq '[Content_Types].xml' -or $_.Extension -eq '.nuspec' } |
        Remove-Item -Force

    if (-not (Test-Path -LiteralPath $Manifest)) {
        throw "PSAppDeployToolkit $Version did not extract correctly to $Target."
    }

    $Actual = (Select-String -LiteralPath $Manifest -Pattern "ModuleVersion\s*=\s*'([^']+)'").Matches[0].Groups[1].Value
    Write-Step "PSAppDeployToolkit $Actual downloaded to $Target" 'OK'
    return $Target
}

function Resolve-IntuneWinAppUtil {
    <#
    .SYNOPSIS
        Returns the path to IntuneWinAppUtil.exe, downloading it on first use.
    .DESCRIPTION
        Resolution order: -IntuneWinAppUtilPath, then the tools cache, then PATH, then a
        download from Microsoft's official repository. The downloaded file is rejected and
        deleted unless it carries a valid Microsoft Authenticode signature.
    #>
    param([string]$Explicit, [string]$ToolsPath, [string]$Version, [switch]$NoDownload)

    if ($Explicit) {
        if (-not (Test-Path -LiteralPath $Explicit -PathType Leaf)) {
            throw "IntuneWinAppUtil.exe not found at: $Explicit"
        }
        return (Resolve-Path -LiteralPath $Explicit).Path
    }

    $Cached = Join-Path -Path $ToolsPath -ChildPath 'IntuneWinAppUtil.exe'
    if (Test-Path -LiteralPath $Cached) {
        Write-Step "IntuneWinAppUtil.exe already cached at $Cached" 'OK'
        return $Cached
    }

    $OnPath = Get-Command -Name 'IntuneWinAppUtil.exe' -ErrorAction SilentlyContinue
    if ($OnPath) {
        Write-Step "IntuneWinAppUtil.exe found on PATH at $($OnPath.Source)" 'OK'
        return $OnPath.Source
    }

    if ($NoDownload) {
        throw "IntuneWinAppUtil.exe is not cached, not on PATH, and -NoDownload was specified. Pass -IntuneWinAppUtilPath."
    }

    Write-Step "Downloading IntuneWinAppUtil.exe $Version from github.com/microsoft..."
    New-Item -Path $ToolsPath -ItemType Directory -Force | Out-Null
    $Uri = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/$Version/IntuneWinAppUtil.exe"

    $OldProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Uri -OutFile $Cached -UseBasicParsing -TimeoutSec 180
    }
    finally {
        $ProgressPreference = $OldProgress
    }

    Assert-MicrosoftSignature -Path $Cached
    Write-Step "IntuneWinAppUtil.exe downloaded to $Cached" 'OK'
    return $Cached
}

function Restore-PackageRoot {
    <#
    .SYNOPSIS
        Stages the toolkit files into an app's package root.
    .DESCRIPTION
        The PSAppDeployToolkit module is roughly 19 MB per package. Committing a copy inside
        every app folder bloats the repository, so this repo ships only the wrapper, the
        payload and the detection script. This function puts the toolkit back, from the
        cached download, producing the same layout New-ADTTemplate would.

        Existing toolkit files are left alone unless -Force is given, so it is safe to run
        repeatedly.
    .OUTPUTS
        'Restored', 'AlreadyPresent' or 'NoWrapper'.
    #>
    param(
        [System.IO.DirectoryInfo]$App,
        [string]$ToolkitPath,
        [switch]$Force
    )

    $Pkg = Join-Path -Path $App.FullName -ChildPath 'AppInstaller'
    $Wrapper = Join-Path -Path $Pkg -ChildPath 'Invoke-AppDeployToolkit.ps1'
    $Module = Join-Path -Path $Pkg -ChildPath 'PSAppDeployToolkit'

    # Only PSADT packages get restored. A Win32 script package has no wrapper and needs none.
    if (-not (Test-Path -LiteralPath $Wrapper)) {
        return 'NoWrapper'
    }

    if ((Test-Path -LiteralPath (Join-Path $Module 'PSAppDeployToolkit.psd1')) -and -not $Force) {
        return 'AlreadyPresent'
    }

    if (Test-Path -LiteralPath $Module) { Remove-Item -LiteralPath $Module -Recurse -Force }

    # The module, minus the parts that are never used at run time.
    Copy-Item -LiteralPath $ToolkitPath -Destination $Module -Recurse -Force
    foreach ($Unused in 'Frontend', 'ADMX') {
        $p = Join-Path -Path $Module -ChildPath $Unused
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    }
    Get-ChildItem -LiteralPath $Module -Filter 'PSGetModuleInfo.xml' -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force

    # The launcher and the extensions module come from the toolkit's v4 frontend.
    Copy-Item -LiteralPath (Join-Path $ToolkitPath 'Frontend\v4\Invoke-AppDeployToolkit.exe') -Destination $Pkg -Force
    $Extensions = Join-Path -Path $Pkg -ChildPath 'PSAppDeployToolkit.Extensions'
    if (Test-Path -LiteralPath $Extensions) { Remove-Item -LiteralPath $Extensions -Recurse -Force }
    Copy-Item -LiteralPath (Join-Path $ToolkitPath 'Frontend\v4\PSAppDeployToolkit.Extensions') -Destination $Pkg -Recurse -Force

    foreach ($Dir in 'Files', 'SupportFiles') {
        $p = Join-Path -Path $Pkg -ChildPath $Dir
        if (-not (Test-Path -LiteralPath $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
    }

    return 'Restored'
}

function Test-PSADTPackage {
    <#
    .SYNOPSIS
        Validates one app's AppInstaller folder as a PSADT package root.
    .OUTPUTS
        'Passed', 'Failed', or 'NotPSADT' when the folder is a plain Win32 script package.
    #>
    param([System.IO.DirectoryInfo]$App)

    $Pkg = Join-Path -Path $App.FullName -ChildPath 'AppInstaller'
    $Wrapper = Join-Path -Path $Pkg -ChildPath 'Invoke-AppDeployToolkit.ps1'
    $Launcher = Join-Path -Path $Pkg -ChildPath 'Invoke-AppDeployToolkit.exe'
    $Module = Join-Path -Path $Pkg -ChildPath 'PSAppDeployToolkit\PSAppDeployToolkit.psd1'
    $FilesDir = Join-Path -Path $Pkg -ChildPath 'Files'

    if (-not (Test-Path -LiteralPath $Pkg)) {
        Write-Step "$($App.Name): no AppInstaller folder - skipped" 'WARN'
        return 'NotPSADT'
    }

    # A plain Win32 script package has no wrapper. That is a valid state, not a failure.
    if (-not (Test-Path -LiteralPath $Wrapper)) {
        Write-Step "$($App.Name): Win32 script package (no PSADT wrapper) - skipped" 'WARN'
        return 'NotPSADT'
    }

    # Each problem carries the remedy with it, so an admin reading the console output does
    # not have to come back to this script to find out what to do about it.
    $Problems = New-Object System.Collections.Generic.List[psobject]
    function Add-Problem {
        param([string]$What, [string]$Fix)
        $Problems.Add([pscustomobject]@{ What = $What; Fix = $Fix })
    }

    # 1. Launcher.
    if (-not (Test-Path -LiteralPath $Launcher)) {
        Add-Problem 'Invoke-AppDeployToolkit.exe is missing - it is the file you name as the Intune setup file' `
            'run this script with -Restore, which stages it along with the toolkit module'
    }

    # 2. Wrapper parses.
    $ParseErrors = $null
    $Tokens = $null
    [System.Management.Automation.Language.Parser]::ParseFile($Wrapper, [ref]$Tokens, [ref]$ParseErrors) | Out-Null
    if ($ParseErrors) {
        Add-Problem "Invoke-AppDeployToolkit.ps1 has $($ParseErrors.Count) syntax error(s), first at line $($ParseErrors[0].Extent.StartLineNumber): $($ParseErrors[0].Message)" `
            'fix the syntax error - on a device this surfaces only as toolkit exit code 60008'
    }

    # 3. Toolkit module.
    if (-not (Test-Path -LiteralPath $Module)) {
        Add-Problem 'PSAppDeployToolkit\PSAppDeployToolkit.psd1 is missing from the package root' `
            'run this script with -Restore to download the toolkit and stage it; Intune deploys to devices that do not have PSADT installed'
    }

    $WrapperText = Get-Content -LiteralPath $Wrapper -Raw

    # 4. Module import path. Test for the corrected guard rather than the absence of the
    #    stock path - the stock path is mentioned in the template's own comments, so a plain
    #    text search for it gives a false positive.
    if ($WrapperText -notmatch '(?m)Test-Path -LiteralPath "\$PSScriptRoot\\PSAppDeployToolkit\\PSAppDeployToolkit\.psd1"') {
        Add-Problem 'the module import does not point at "$PSScriptRoot\PSAppDeployToolkit\..." - the package will fail on any device without PSADT installed, with exit code 60008' `
            'in Invoke-AppDeployToolkit.ps1 replace $PSScriptRoot\..\..\..\PSAppDeployToolkit with $PSScriptRoot\PSAppDeployToolkit (both the Test-Path and the Import-Module lines)'
    }

    # 5. The wrapper's payload script exists.
    $Match = [regex]::Match($WrapperText, "(?m)^\s*\`$InstallerScriptName\s*=\s*'([^']+)'")
    if (-not $Match.Success) {
        Write-Step "$($App.Name): no `$InstallerScriptName variable found in the wrapper - skipping that check" 'WARN'
    }
    else {
        $Expected = $Match.Groups[1].Value
        if (Test-Path -LiteralPath (Join-Path $FilesDir $Expected)) {
            Write-Step "$($App.Name): wrapper targets [$Expected] - present in Files\" 'OK'
        }
        else {
            Add-Problem "the wrapper targets [$Expected] but that file is not in Files\" `
                "either put [$Expected] in Files\, or update `$InstallerScriptName at the top of Invoke-AppDeployToolkit.ps1 to match the file that is there"
        }
    }

    # 6. Payload sanity.
    if (-not (Test-Path -LiteralPath $FilesDir)) {
        Add-Problem 'Files\ is missing' `
            'create Files\ in the package root and put the installer script and its config files in it'
    }
    else {
        $Payload = @(Get-ChildItem -LiteralPath $FilesDir -File)
        if ($Payload.Count -eq 0) {
            Add-Problem 'Files\ is empty' `
                'copy the installer script and its config files into Files\ - this is the payload the wrapper runs'
        }
        else {
            Write-Step "$($App.Name): payload [$(($Payload.Name) -join ', ')]" 'OK'
        }

        if (Test-Path -LiteralPath (Join-Path $FilesDir 'PSAppDeployToolkit')) {
            Add-Problem 'Files\ contains a nested PSAppDeployToolkit folder' `
                'the toolkit module belongs in the package root, not inside Files\ - delete the nested copy'
        }
    }

    # 7. Path length.
    $Longest = (Get-ChildItem -LiteralPath $Pkg -Recurse -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.FullName.Length } | Measure-Object -Maximum).Maximum
    if ($Longest -gt 240) {
        Add-Problem "longest path inside the package is $Longest characters" `
            'move the repository closer to the drive root - past MAX_PATH the toolkit DLLs fail to load and you get exit code 60008'
    }

    # 8. Toolkit file signatures. PSAppDeployToolkit ships config.psd1, strings.psd1 and the
    #    module scripts Authenticode-signed. Hand-editing any of them (e.g. to point Assets at
    #    a custom banner) invalidates the signature, and Initialize-ADTModule refuses to start
    #    the session - on a device this is exit code 60008 with no useful message, and it can
    #    strand the toolkit mid-dialog if the signature check fails after the UI already spawned.
    $SignedFiles = @(Get-ChildItem -LiteralPath $Module.Replace('\PSAppDeployToolkit.psd1', '') -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.psd1', '.psm1', '.ps1' })
    foreach ($SignedFile in $SignedFiles) {
        $Signature = Get-AuthenticodeSignature -LiteralPath $SignedFile.FullName
        if ($Signature.Status -ne 'Valid' -and $Signature.Status -ne 'NotSigned') {
            Add-Problem "$($SignedFile.FullName.Substring($Pkg.Length + 1)) has an invalid signature ($($Signature.Status)) - it was edited after being signed" `
                'revert the file to the shipped PSAppDeployToolkit content; swap branding by replacing the asset files under Assets\ with the same filenames instead of editing Config\config.psd1'
        }
    }

    if ($Problems.Count -eq 0) {
        $SizeMB = [math]::Round((Get-ChildItem -LiteralPath $Pkg -Recurse -File | Measure-Object Length -Sum).Sum / 1MB, 1)
        Write-Step "$($App.Name): PSADT package OK ($SizeMB MB, longest path $Longest chars)" 'OK'
        return 'Passed'
    }

    foreach ($Problem in $Problems) {
        Write-Step "$($App.Name): $($Problem.What)" 'FAIL'
        Write-Host ('         Fix: {0}' -f $Problem.Fix) -ForegroundColor DarkYellow
    }
    return 'Failed'
}

function Build-IntuneWin {
    param(
        [System.IO.DirectoryInfo]$App,
        [string]$UtilPath
    )

    $SourceFolder = Join-Path -Path $App.FullName -ChildPath 'AppInstaller'
    $SetupFile = 'Invoke-AppDeployToolkit.exe'
    $OutputFolder = Join-Path -Path $App.FullName -ChildPath 'Package'

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    Write-Step "$($App.Name): building .intunewin from [$SetupFile]..."
    $Arguments = @('-c', "`"$SourceFolder`"", '-s', "`"$SetupFile`"", '-o', "`"$OutputFolder`"", '-q')
    $Process = Start-Process -FilePath $UtilPath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow

    if ($Process.ExitCode -eq 0) {
        $Built = Get-ChildItem -LiteralPath $OutputFolder -Filter '*.intunewin' |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Step "$($App.Name): built $($Built.Name) ($([math]::Round($Built.Length / 1MB, 1)) MB)" 'OK'
    }
    else {
        Write-Step "$($App.Name): IntuneWinAppUtil failed with exit code $($Process.ExitCode)" 'FAIL'
    }
}


# ---------------------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------------------

if ($SolutionRoot) {
    if (-not (Test-Path -LiteralPath $SolutionRoot -PathType Container)) {
        Write-Step "SolutionRoot not found: $SolutionRoot" 'FAIL'
        exit 1
    }
    $SolutionRoot = (Resolve-Path -LiteralPath $SolutionRoot).Path
}
else {
    $SolutionRoot = Find-SolutionRoot -StartPath $PSScriptRoot
}

if ($PackagePath) {
    $Apps = @($PackagePath | ForEach-Object { Get-Item -LiteralPath $_ })
}
elseif ($SolutionRoot) {
    $Apps = @(Get-ChildItem -LiteralPath $SolutionRoot -Directory | Where-Object {
            Test-Path -LiteralPath (Join-Path $_.FullName 'AppInstaller')
        })
}
else {
    $Apps = @()
}

Write-Host ''
Write-Host 'Build-PSADTPackage 3.0 - validates PSADT v4 package folders before packaging.' -ForegroundColor White
Write-Host 'Read-only unless -Restore is used. First run after cloning: -Restore. Help: -?' -ForegroundColor DarkGray
Write-Host ''

if ($Apps.Count -eq 0) {
    Write-Step 'No app folders were found.' 'FAIL'
    Write-Step "An app folder is any folder containing an 'AppInstaller' subfolder." 'INFO'
    Write-Step "Searched upward from: $PSScriptRoot" 'INFO'
    Write-Step 'Point at them explicitly with -SolutionRoot <path> or -PackagePath <path>.' 'INFO'
    exit 1
}

if (-not $ToolsPath) { $ToolsPath = Join-Path -Path $SolutionRoot -ChildPath '.tools' }

Write-Step "Solution root: $SolutionRoot"
Write-Step "Apps found   : $($Apps.Name -join ', ')"
Write-Step "Tools cache  : $ToolsPath"
Write-Host ''

# --- Restore: put the toolkit into each package root -----------------------------------
# The module is not committed to the repository, so a fresh clone needs this once.
if ($Restore) {
    $ToolkitPath = Get-PSADTToolkit -ToolsPath $ToolsPath -Version $PSADTVersion -NoDownload:$NoDownload
    Write-Host ''

    foreach ($App in $Apps) {
        switch (Restore-PackageRoot -App $App -ToolkitPath $ToolkitPath -Force:$Force) {
            'Restored' { Write-Step "$($App.Name): toolkit staged into AppInstaller\" 'OK' }
            'AlreadyPresent' { Write-Step "$($App.Name): toolkit already present - use -Force to replace it" 'WARN' }
            'NoWrapper' { Write-Step "$($App.Name): Win32 script package - nothing to restore" 'WARN' }
        }
    }
    Write-Host ''
}

$UtilPath = $null
if ($BuildIntuneWin) {
    $UtilPath = Resolve-IntuneWinAppUtil -Explicit $IntuneWinAppUtilPath -ToolsPath $ToolsPath `
        -Version $IntuneWinAppUtilVersion -NoDownload:$NoDownload
    Write-Host ''
}

$Failures = 0

foreach ($App in $Apps) {
    $Result = Test-PSADTPackage -App $App

    if ($Result -eq 'Failed') { $Failures++; continue }
    if ($Result -eq 'NotPSADT') { continue }

    if ($BuildIntuneWin) {
        Build-IntuneWin -App $App -UtilPath $UtilPath
    }
    Write-Host ''
}

Write-Host ''
if ($Failures -gt 0) {
    Write-Step "$Failures package(s) need attention - see the Fix lines above. Nothing was built." 'FAIL'
    exit 1
}

if ($BuildIntuneWin) {
    Write-Step 'All packages passed and the .intunewin files are in each app''s Package folder.' 'OK'
}
else {
    Write-Step 'All packages passed. Add -BuildIntuneWin to produce the .intunewin files.' 'OK'
}
exit 0
