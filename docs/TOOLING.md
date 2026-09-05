# Tooling reference

Everything in `Install-Tools\` is a workstation tool. None of it is deployed and none of it
belongs inside a package you upload to Intune.

## Build-PSADTPackage.ps1

### The problem it solves

A PSADT package that is assembled slightly wrong still looks perfectly fine on your
packaging workstation. It builds, the `.intunewin` uploads, Intune accepts it — and then it
fails on every device with **exit code 60008** and no useful message in the portal.

The most common causes: the wrapper that ships in the toolkit's `Frontend\v4` folder imports
its module from `$PSScriptRoot\..\..\..\PSAppDeployToolkit`, a path that only resolves while
the file is sitting three levels inside the extracted toolkit — copy it into a package root
unchanged and the import silently falls through to `PSModulePath`, which is empty on any
device without PSADT already installed. Or a signed toolkit file (typically
`Config\config.psd1`) was hand-edited after being signed, which makes `Open-ADTSession`
refuse to start on every device.

This script catches those and six other assembly mistakes in about a second.

### Quick start

```powershell
# Check everything. Read-only, nothing is written.
.\Build-PSADTPackage.ps1

# Check one app
.\Build-PSADTPackage.ps1 -PackagePath ..\Microsoft-365-Apps-Business

# Check, then build the .intunewin into that app's Package\ folder
.\Build-PSADTPackage.ps1 -PackagePath ..\Microsoft-365-Apps-Business `
    -BuildIntuneWin -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe

# Full help
Get-Help .\Build-PSADTPackage.ps1 -Full
```

Requires Windows PowerShell 5.1 or PowerShell 7+. **No administrator rights.**
`-BuildIntuneWin` additionally needs
[IntuneWinAppUtil.exe](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool).

### Expected folder layout

The script looks for app folders — any folder containing an `AppInstaller` subfolder — and
treats `AppInstaller` itself as the PSADT package root. See
[Packaging](PACKAGING.md#folder-layout) for the full tree.

It finds the solution root by walking up from its own location, so it works whether you keep
it beside the app folders or in a subfolder like this one. Override with `-SolutionRoot`.

An `AppInstaller` with no `Invoke-AppDeployToolkit.ps1` is reported as **skipped** — that is
a plain Win32 script package, which is a valid state, not a failure.

### What it checks

| # | Check | Why it matters |
|---|---|---|
| 1 | `Invoke-AppDeployToolkit.exe` present | It is the setup file you name in Intune |
| 2 | Wrapper has no PowerShell syntax errors | On a device this appears only as exit code 60008 |
| 3 | Toolkit module present in the package root | Target devices have no PSADT installed |
| 4 | Module import points at `$PSScriptRoot\PSAppDeployToolkit\...` | **The 60008 trap described above** |
| 5 | The script named by `$InstallerScriptName` exists in `Files\` | Catches bumping the installer version and forgetting the wrapper |
| 6 | `Files\` is not empty and has no nested package folder | An empty `Files\` means the payload was never copied in |
| 7 | No path over 240 characters | Past MAX_PATH the toolkit's own DLLs fail to load |
| 8 | Every `.psd1`/`.psm1`/`.ps1` under `PSAppDeployToolkit\` still has a valid Authenticode signature | A hand edit (e.g. to `Config\config.psd1`) invalidates it and `Open-ADTSession` refuses to start on every device |

### Reading the output

```text
[OK  ]  check passed
[WARN]  informational, does not fail the run
[FAIL]  the package would not work as-is
        Fix: <what to do about it>
```

Every `[FAIL]` is followed by a `Fix:` line, so you do not need to come back here.

Example of a broken package:

```text
[FAIL] BrokenApp: the module import does not point at "$PSScriptRoot\PSAppDeployToolkit\..." -
       the package will fail on any device without PSADT installed, with exit code 60008
         Fix: in Invoke-AppDeployToolkit.ps1 replace $PSScriptRoot\..\..\..\PSAppDeployToolkit
              with $PSScriptRoot\PSAppDeployToolkit (both the Test-Path and the Import-Module lines)
```

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Every package checked passed |
| `1` | At least one package failed, or no app folders were found |

Nothing is built when a package fails, so it is safe to use as a CI gate:

```yaml
- run: pwsh -File Install-Tools/Build-PSADTPackage.ps1
```

### Adapting it to your own repo

- Different package root folder name? Change the `'AppInstaller'` literal in
  `Find-SolutionRoot` and `Test-PSADTPackage`.
- Not using an `$InstallerScriptName` variable in your wrapper? Check 5 is skipped
  automatically with a `[WARN]`, and the rest still apply.
- Different PSADT version? The checks are version-agnostic; only the folder names matter.

## Invoke-AppDeployToolkit_M365-Template.ps1

A master copy of the PSADT wrapper. Copy it into a package root, rename it to
`Invoke-AppDeployToolkit.ps1`, and it is ready to use — no other editing for the Business
package; two lines for Enterprise.

It is the stock PSADT v4 template plus four changes:

- Module import corrected to `$PSScriptRoot\PSAppDeployToolkit\...` (the 60008 trap).
- Adds a `-ProductID` parameter, excluded from `Open-ADTSession` via
  `Get-ADTBoundParametersAndDefaultValues -Exclude ProductID`, because the session rejects
  parameters it does not own.
- Launches the wrapped installer through native-bitness PowerShell.
- Uses `Start-ADTProcess -IgnoreExitCodes '*' -PassThru` so the real installer exit code
  reaches Intune instead of being masked as the generic toolkit failure `60001`.

The end-user behaviour it implements:

| Situation | What happens |
|---|---|
| No Office app running | No dialog. Install proceeds immediately. |
| Office app running, user logged on | Countdown dialog listing the apps, with a **Close Programs** button and no **Postpone**. Closing the app yourself (not clicking the button) also dismisses the dialog within a couple of seconds — it isn't waiting on the button specifically. At the end of the countdown the apps are closed for the user. |
| Autopilot / ESP / no user | PSADT runs Silent and closes anything running without prompting. |

Deferral is deliberately not enabled: PSADT returns `60012` when a user defers, which Intune
treats as a failure unless mapped to Retry, and a deferred install during Autopilot would
strand the device.

## Other notes

- `Build-PSADTPackage.ps1` is a **workstation tool**. It never goes inside a package and is
  never deployed. It copies nothing by default — only `-Restore` writes into a package.

- Do not edit the cached toolkit under `.tools\`. It is a download, and `-Restore -Force`
  will replace it. Customisation belongs in the package's own `Invoke-AppDeployToolkit.ps1`
  (behavior) or in `Assets\*.png` file content under the same filenames (branding) — never in
  `Config\config.psd1`'s paths, which breaks its signature. See
  [Packaging](PACKAGING.md#folder-layout).

- **Keep the folder path short.** PSADT's `Strings` and `lib` trees are deep; the module files
  already reach about 200 characters from the solution root. During development a copy under a
  long temp path exceeded 260 characters and PSADT failed to load its DLLs with exit code
  60008. Intune extracts to `C:\Windows\IMECache\`, so this only bites locally.

- `FORCEAPPSHUTDOWN` was removed from `Configuration.xml` so users get the PSADT countdown
  instead of ODT killing Office without warning. It is deliberately still present in
  `Uninstall.xml`, so a removal cannot be blocked by an open app.

- The language packs and proofing tools packages are Win32 script packages only. Their
  `AppInstaller` folders hold just the script and the XML files, so `Build-PSADTPackage.ps1`
  reports them as `Win32 script package (no PSADT wrapper) - skipped`. That is expected and
  not an error. To give them a PSADT package later, follow [Packaging](PACKAGING.md) — but
  their wrappers need a `-LanguageID` parameter instead of `-ProductID`, so the M365 template
  needs editing rather than copying as-is.

- **Resolved:** *"uninstall does nothing unless the device is restarted"*. Cause was `setup.exe`
  being started while the Click-to-Run engine was still executing the install scenario.
  Fixed in v2.5 by `Wait-OfficeEngineIdle`. Verify with the
  [testing checklist](TESTING.md#4-uninstall-immediately-after-install-without-a-reboot).

- **Resolved:** *"the script uninstalls other Office versions"*. `Uninstall.xml` targets a single
  Product ID and v2.5 checks `Test-ProductInstalled` first, so a product that is not present is
  a no-op that exits 0. Verify with the
  [testing checklist](TESTING.md#3-product-isolation-this-was-a-real-false-positive).

- **Resolved:** *"the close-apps popup never goes away" / "uninstall does nothing with an app
  open"*. Both were `config.psd1`'s Authenticode signature being broken by a hand edit, not a
  defect in the dialog or the uninstall logic. See
  [Troubleshooting](TROUBLESHOOTING.md#the-uninstall-does-nothing-when-an-office-app-is-open--the-close-apps-popup-never-goes-away).
