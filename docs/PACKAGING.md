# Packaging

How the `AppInstaller` folders are laid out, how to (re)build a PSADT package from
scratch, and how to produce the `.intunewin` file. Most of the time you only need the
one-liner in the main [README](../README.md#quick-start) — this page is for building a
package by hand, converting between packaging methods, or understanding what
`Build-PSADTPackage.ps1` assembled for you.

## Folder layout

In this solution the **`AppInstaller` folder IS the PSADT package root**. The toolkit, the
launcher, the wrapper and the payload all live inside it, and that whole folder is what
`IntuneWinAppUtil` wraps.

```text
<App>\
  AppInstaller\                      <- PSADT package root, wrap THIS
    Invoke-AppDeployToolkit.ps1        wrapper (customised)
    Invoke-AppDeployToolkit.exe        Intune setup file
    PSAppDeployToolkit\                toolkit module (no Frontend, no ADMX)
    PSAppDeployToolkit.Extensions\
    Files\                             installer script + XML files
    SupportFiles\                      extra files not read by any script (see note below)
  Detection\                         <- uploaded separately to Intune
  Package\                           <- .intunewin output
```

Current state:

| App | Packaging |
|---|---|
| `Microsoft-365-Apps-Business` | PSADT package |
| `Microsoft-365-Apps-Enterprise` | PSADT package |
| `Deploy-Microsoft-365-LanguagePacks` | Win32 script package (no PSADT files) |
| `Deploy-Microsoft-365-ProofingTools` | Win32 script package (no PSADT files) |

Because the payload exists only once, the two packaging methods cannot drift apart.

> **`SupportFiles\` is not wired up.** The M365 wrapper never reads
> `$adtSession.DirSupportFiles`, so anything placed there (e.g. `AppLogo.png`) travels inside
> the `.intunewin` but has no effect on install/uninstall/repair. It's a fine place to keep a
> reference copy of the icon you upload by hand to Intune's **App information → Logo** field —
> just don't expect the package to use it automatically. To actually change what a PSADT
> dialog shows, replace the image *content* of the same-named files under
> `PSAppDeployToolkit\Assets\` (`AppIcon.png`, `Banner.Classic.png`) instead of adding new
> filenames — see the signature warning below.

> **Do not edit files inside `PSAppDeployToolkit\`.** `Config\config.psd1`, `Strings\*.psd1`
> and the module scripts are Authenticode-signed by the toolkit's publisher. Changing so much
> as one character (e.g. pointing `Assets.Banner` at a new filename) invalidates the signature,
> and `Open-ADTSession` refuses to start — on a device this is exit code 60008 with no useful
> message, and it can strand the toolkit mid-dialog if the check fails after the UI already
> spawned. `Build-PSADTPackage.ps1` checks every `.psd1`/`.psm1`/`.ps1` signature and fails the
> build if one doesn't match. Swap branding by overwriting the *image content* of
> `Assets\AppIcon.png` / `Assets\Banner.Classic.png` — those aren't part of the signed
> manifest — and leave `config.psd1` untouched.

### Converting a PSADT package back to a Win32 script package

Delete the PSADT files and lift the payload back up one level:

```powershell
$Pkg = "<App>\AppInstaller"
Remove-Item -LiteralPath "$Pkg\PSAppDeployToolkit" -Recurse -Force
Remove-Item -LiteralPath "$Pkg\PSAppDeployToolkit.Extensions" -Recurse -Force
Remove-Item -LiteralPath "$Pkg\Invoke-AppDeployToolkit.ps1" -Force
Remove-Item -LiteralPath "$Pkg\Invoke-AppDeployToolkit.exe" -Force
Remove-Item -LiteralPath "$Pkg\SupportFiles" -Recurse -Force
Get-ChildItem -LiteralPath "$Pkg\Files" -File |
    ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $Pkg -Force }
Remove-Item -LiteralPath "$Pkg\Files" -Recurse -Force
```

`AppInstaller\` then holds just the installer script and the XML files, and the Intune
command reverts to the `powershell.exe` form in
[Intune setup](INTUNE-SETUP.md#microsoft-365-apps-business--enterprise).

## Build a PSADT package from the shared toolkit copy

A PSADT package root must be self-contained. `-Restore` does this for you; the steps below
are the manual equivalent, using the cached toolkit under `.tools\`.

Run this once per app. Example uses `Microsoft-365-Apps-Business`.

### The short way: New-ADTTemplate

PSAppDeployToolkit ships a cmdlet that creates a correct package root for you. Use it unless
you have a reason not to:

```powershell
Import-Module "<toolkit master copy>\PSAppDeployToolkit.psd1"
New-ADTTemplate -Destination "<App>" -Name AppInstaller -Version 4
```

That produces `<App>\AppInstaller\` containing `Invoke-AppDeployToolkit.ps1` and `.exe`, the
`PSAppDeployToolkit` module, `PSAppDeployToolkit.Extensions`, `Files\`, `SupportFiles\` and
the optional `Assets\`, `Config\` and `Strings\` override folders — with the module import
path already correct.

Then:

1. Copy your installer script and config files into `AppInstaller\Files\`.
2. Replace `AppInstaller\Invoke-AppDeployToolkit.ps1` with
   `Install-Tools\Invoke-AppDeployToolkit_M365-Template.ps1`, which already contains the
   Microsoft 365 deployment logic.
3. Validate with `Build-PSADTPackage.ps1` (below).

The manual steps below do the same thing by hand. They are worth reading once because they
explain what each piece of the package root is for, but for routine work `New-ADTTemplate` is
fewer steps and cannot get the paths wrong.

### Manual assembly

**1. Set the paths**

```powershell
$Root = "C:\Microsoft-365\Microsoft Intune\Application-Deployment"
$Tk   = "$Root\Deploy-Microsoft-365-Apps\.tools\PSAppDeployToolkit-4.1.8"
$App  = "$Root\Deploy-Microsoft-365-Apps\Microsoft-365-Apps-Business"
$Pkg  = "$App\AppInstaller"
```

**2. Move the existing payload into `Files\`**

Start from an `AppInstaller` folder that holds the installer script and the XML files.

```powershell
New-Item -ItemType Directory -Path "$Pkg\Files" -Force | Out-Null
New-Item -ItemType Directory -Path "$Pkg\SupportFiles" -Force | Out-Null
Get-ChildItem -LiteralPath $Pkg -File |
    ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination "$Pkg\Files" -Force }
```

**3. Copy the toolkit module**

```powershell
Copy-Item -LiteralPath $Tk -Destination $Pkg -Recurse -Force
Remove-Item -LiteralPath "$Pkg\PSAppDeployToolkit\Frontend" -Recurse -Force
Remove-Item -LiteralPath "$Pkg\PSAppDeployToolkit\ADMX" -Recurse -Force
Get-ChildItem -LiteralPath "$Pkg\PSAppDeployToolkit" -Filter PSGetModuleInfo.xml -Force |
    Remove-Item -Force
```

`Frontend` and `ADMX` are removed because they are not used at run time. Everything else in
the module folder is required (`Config`, `Strings`, `Assets`, `lib`). Result is about 19 MB.

**4. Copy the launcher and the extensions module**

```powershell
Copy-Item -LiteralPath "$Tk\Frontend\v4\Invoke-AppDeployToolkit.exe" -Destination $Pkg -Force
Copy-Item -LiteralPath "$Tk\Frontend\v4\PSAppDeployToolkit.Extensions" -Destination $Pkg -Recurse -Force
```

> Do **not** copy `Frontend\v4\Invoke-AppDeployToolkit.ps1` — use the template instead (step 5).

**5. Add the deployment script**

For the M365 Apps packages, copy the ready-made wrapper:

```powershell
Copy-Item -LiteralPath "$Root\Deploy-Microsoft-365-Apps\Install-Tools\Invoke-AppDeployToolkit_M365-Template.ps1" `
          -Destination "$Pkg\Invoke-AppDeployToolkit.ps1" -Force
```

For Enterprise, then change two things in `$Pkg\Invoke-AppDeployToolkit.ps1`:

```powershell
[System.String]$ProductID = 'O365ProPlusRetail'
AppName = 'Microsoft 365 Apps for Enterprise'
```

#### If you start from the stock `Frontend\v4\Invoke-AppDeployToolkit.ps1` instead

> **PSADT is not wrong here — this is a hand-copying mistake.** The stock file lives at
> `<toolkit>\PSAppDeployToolkit\Frontend\v4\`, and from there `..\..\..\PSAppDeployToolkit`
> resolves correctly, so the template can be run in place for testing. `New-ADTTemplate`
> rewrites the path to `$PSScriptRoot\PSAppDeployToolkit` when it generates a package.
>
> If you copy the file by hand instead of using `New-ADTTemplate` you skip that rewrite. The
> import then silently falls through to `PSModulePath`, which is empty on a managed device,
> and every deployment fails with exit code **60008**. Make this edit, or use
> `New-ADTTemplate` and avoid the problem entirely.

Find (three occurrences):

```text
"$PSScriptRoot\..\..\..\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
"$PSScriptRoot\..\..\..\PSAppDeployToolkit"
```

Replace with:

```text
"$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
"$PSScriptRoot\PSAppDeployToolkit"
```

One-liner for that edit:

```powershell
$f = "$Pkg\Invoke-AppDeployToolkit.ps1"
(Get-Content -LiteralPath $f -Raw).Replace('$PSScriptRoot\..\..\..\PSAppDeployToolkit','$PSScriptRoot\PSAppDeployToolkit') |
    Set-Content -LiteralPath $f -Encoding UTF8
```

Then add the deployment logic inside `Install-ADTDeployment` / `Uninstall-ADTDeployment`.
The minimum that makes the installer script run:

```powershell
$script = Join-Path $adtSession.DirFiles 'Install-Microsoft-365-Apps_v2.5.ps1'
$result = Start-ADTProcess -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
              -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$script`" -Mode Install -ProductID O365BusinessRetail" `
              -WorkingDirectory $adtSession.DirFiles -CreateNoWindow -IgnoreExitCodes '*' -PassThru
if ($result.ExitCode -ne 0) { Close-ADTSession -ExitCode $result.ExitCode }
```

`-IgnoreExitCodes '*'` matters: without it a non-zero ODT code throws, and Intune sees the
generic toolkit failure **60001** instead of the real reason.

If you add your own parameter to the `param()` block (like `-ProductID`), it must also be
excluded from the session, because `Open-ADTSession` rejects parameters it does not own:

```powershell
$iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation -Exclude ProductID
```

**6. Validate the package**

```powershell
.\Build-PSADTPackage.ps1 -PackagePath "$App"
```

Checks the launcher, the wrapper syntax, the module import path, that the script named in
`$InstallerScriptName` exists in `Files\`, the toolkit file signatures, and the longest path.
Full list of checks: [Tooling reference](TOOLING.md#what-it-checks). Expected output:

```text
[OK  ] Microsoft-365-Apps-Business: wrapper targets [Install-Microsoft-365-Apps_v2.5.ps1] - present in Files\
[OK  ] Microsoft-365-Apps-Business: payload [Configuration.xml, Install-Microsoft-365-Apps_v2.5.ps1, Uninstall.xml]
[OK  ] Microsoft-365-Apps-Business: PSADT package OK (19.2 MB, longest path 201 chars)
```

**7. Final layout to verify before packaging**

```text
AppInstaller\
  Invoke-AppDeployToolkit.ps1        <- edited wrapper
  Invoke-AppDeployToolkit.exe        <- Intune setup file
  PSAppDeployToolkit\                <- module, no Frontend, no ADMX
  PSAppDeployToolkit.Extensions\
  Files\ Configuration.xml, Install-Microsoft-365-Apps_v2.5.ps1, Uninstall.xml
  SupportFiles\
```

Manual equivalent of the validator, if you prefer to check by hand:

```powershell
Test-Path "$Pkg\PSAppDeployToolkit\PSAppDeployToolkit.psd1"   # must be True
Test-Path "$Pkg\Invoke-AppDeployToolkit.exe"                  # must be True
Get-ChildItem "$Pkg\Files"

Select-String -LiteralPath "$Pkg\Invoke-AppDeployToolkit.ps1" `
    -Pattern 'Test-Path -LiteralPath "\$PSScriptRoot\\PSAppDeployToolkit\\PSAppDeployToolkit\.psd1"'
# must return exactly ONE line. If it returns nothing, step 5 was not applied and
# the package will fail on a clean device with exit code 60008.

Get-ChildItem -LiteralPath "$Pkg\PSAppDeployToolkit" -Recurse -File |
    Where-Object Extension -in '.psd1', '.psm1', '.ps1' |
    Get-AuthenticodeSignature | Where-Object Status -ne 'Valid'
# must return nothing. Anything listed here was hand-edited after being signed and
# Open-ADTSession will refuse to start (exit code 60008).
```

> Do **not** test for the *absence* of `..\..\..\PSAppDeployToolkit` — the template mentions
> that path in its own comments, so a plain text search gives a false positive.

## Create the .intunewin

PSADT package (Business, Enterprise):

```text
IntuneWinAppUtil.exe -c "<App>\AppInstaller" -s "Invoke-AppDeployToolkit.exe" -o "<App>\Package" -q
```

Win32 script package (LanguagePacks, ProofingTools):

```text
IntuneWinAppUtil.exe -c "<App>\AppInstaller" -s "Install-LanguagePacks_v3.0.ps1" -o "<App>\Package" -q
```

The `-c` folder is the same in both cases; only `-s` changes, because `AppInstaller` is
either a PSADT package root or a plain source folder depending on what it contains.

> Never point `-c` at a folder that contains the `Package` output folder.

The validator can also run IntuneWinAppUtil for you:

```powershell
.\Build-PSADTPackage.ps1 -PackagePath "$App" -BuildIntuneWin `
    -IntuneWinAppUtilPath "$Root\Deploy-Microsoft-365-Apps\.tools\IntuneWinAppUtil.exe"
```

> **The `.intunewin` files are build output, not source.** They're gitignored
> (`*/Package/*.intunewin`) — rebuild them locally with the command above whenever
> `AppInstaller\` changes, rather than committing them to the repo.
