# Deploy Microsoft 365 Apps with Intune

Deploy **Microsoft 365 Apps for Enterprise** or **Microsoft 365 Apps for Business** from
Intune as a Win32 app, optimised for Autopilot ESP — packaged either as a plain Win32 script
package **or** with **PSAppDeployToolkit**, from the same installer script and the same XML
files. Language packs and proofing tools ship as separate apps in the same repository.

> **Validate and test** both the scripts and the configuration in a controlled test
> environment before deploying to production — see the
> [validation checklist](docs/TESTING.md).

## How it works

The installer downloads the latest Office Deployment Tool from Microsoft's evergreen CDN on
every run and drives `setup.exe` with the supplied `Configuration.xml` or `Uninstall.xml`, so
the package itself stays tiny and never carries a stale Office build. Office is installed in
the **Dutch and English UI languages alongside the OS display language**, so a device keeps
working whichever language Windows is running in.

- Optimised for **Autopilot ESP** — a small package that streams the bits from Microsoft's CDN
- One script handles **Install**, **Uninstall** and **Repair** through the `-Mode` parameter
- Verifies Microsoft's Authenticode signature on `setup.exe` before running it
- Supports **O365ProPlusRetail** and **O365BusinessRetail** through `-ProductID`
- Removes **only** the product it was told to remove; other Office versions are untouched
- Runs 64-bit even though the Intune Management Extension is a 32-bit host
- Waits for the Click-to-Run engine to go idle, so an uninstall issued straight after an
  install works **without a reboot**
- Detection scripts match one **exact** product ID and verify the languages actually landed
- Packages with **PSAppDeployToolkit v4** for a user-facing close-apps countdown, or as a
  plain Win32 script package — your choice, same source files

> **ARM64.** `Configuration.xml` sets `OfficeClientEdition="64"`. On an ARM64 device the
> Office Deployment Tool serves the ARM64 build automatically, so one package covers AMD64
> and ARM64.

## Repository layout

| App folder | Installer script | Detection script | Packaging |
|---|---|---|---|
| `Microsoft-365-Apps-Business` | `Install-Microsoft-365-Apps_v2.5.ps1` | `Detect-Microsoft-365-Apps_v2.0.ps1` | PSADT |
| `Microsoft-365-Apps-Enterprise` | `Install-Microsoft-365-Apps_v2.5.ps1` | `Detect-Microsoft-365-Apps_v2.0.ps1` | PSADT |
| `Deploy-Microsoft-365-LanguagePacks` | `Install-LanguagePacks_v3.0.ps1` | `Detect-LanguagePacks_v3.0.ps1` | Win32 script |
| `Deploy-Microsoft-365-ProofingTools` | `Install-ProofingTools_v3.0.ps1` | `Detect-ProofingTools_v3.0.ps1` | Win32 script |

`Install-Tools\` holds workstation-only tooling (never deployed) — see the
[tooling reference](docs/TOOLING.md).

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+ on the packaging workstation. No admin rights
  needed to build; admin **is** needed to run an installer manually.
- Internet access on the first run, to download PSAppDeployToolkit and IntuneWinAppUtil.exe
  (cached under `.tools\` afterwards — see Quick start).
- A test device you are willing to reset, for the [validation checklist](docs/TESTING.md).

> **Keep the repo path short.** PSADT's `Strings` and `lib` trees are deep and already reach
> ~200 characters from the repo root. Past MAX_PATH the toolkit's DLLs fail to load and every
> deployment returns exit code 60008 — see [Troubleshooting](docs/TROUBLESHOOTING.md).

## Quick start

After cloning, one command fetches the tools, assembles the packages and builds the
`.intunewin` files:

```powershell
cd Install-Tools
.\Build-PSADTPackage.ps1 -Restore -BuildIntuneWin
```

That downloads **PSAppDeployToolkit** and **IntuneWinAppUtil.exe** (verifying its Microsoft
signature), stages the toolkit into each package root, validates every package, and writes
`<App>\Package\Invoke-AppDeployToolkit.intunewin`.

Then upload using the settings in [Intune setup](docs/INTUNE-SETUP.md), and work through the
[validation checklist](docs/TESTING.md) on a test device before broad deployment.

Later runs need no switches, and nothing is downloaded again:

```powershell
.\Build-PSADTPackage.ps1                      # validate only, read-only
.\Build-PSADTPackage.ps1 -BuildIntuneWin      # validate and rebuild the .intunewin files
.\Build-PSADTPackage.ps1 -Restore -Force      # move the packages onto a new PSADT version
```

Air-gapped workstation:

```powershell
.\Build-PSADTPackage.ps1 -Restore -BuildIntuneWin -NoDownload `
    -ToolsPath D:\Offline\Tools -IntuneWinAppUtilPath D:\Offline\IntuneWinAppUtil.exe
```

> **What is not in this repository.** The PSAppDeployToolkit module (~19 MB per package) and
> every `.intunewin` are gitignored build output — `-Restore` / `-BuildIntuneWin` regenerate
> them locally. That's why a fresh clone is under 1 MB, and why there's nothing to commit
> after a rebuild.

## Documentation

| Page | What's in it |
|---|---|
| [Packaging](docs/PACKAGING.md) | Package folder layout, building a PSADT package by hand, converting to/from a plain Win32 script package, creating the `.intunewin` |
| [Intune setup](docs/INTUNE-SETUP.md) | Exact install/uninstall commands, detection rule, return codes, and what the wrapper runs underneath |
| [Testing](docs/TESTING.md) | The validation checklist to run on a test device before production rollout |
| [Troubleshooting](docs/TROUBLESHOOTING.md) | Log locations, manual test commands, exit codes, and fixes for the failures you'll actually hit |
| [Tooling reference](docs/TOOLING.md) | What `Build-PSADTPackage.ps1` checks and why, plus notes on the wrapper template |

## Support files and branding

`AppInstaller\SupportFiles\` (e.g. `AppLogo.png`) isn't read by any script — it's a place to
keep a reference copy of the icon you upload by hand to Intune's **App information → Logo**
field. To change what a PSADT dialog itself shows, overwrite the image content of
`PSAppDeployToolkit\Assets\AppIcon.png` / `Banner.Classic.png` under the same filenames —
**never** edit paths in `PSAppDeployToolkit\Config\config.psd1`, which is Authenticode-signed
and will refuse to load if changed. Details in [Packaging](docs/PACKAGING.md#folder-layout).
