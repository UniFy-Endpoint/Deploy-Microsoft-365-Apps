# Logs, manual testing, and troubleshooting

## Log locations

| What | Path |
|---|---|
| ODT installer | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log` |
| Language packs | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\M365LanguagePackSetup.log` |
| Proofing tools | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\M365ProofingToolsSetup.log` |
| Detection | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Detect-*.log` |
| PSADT wrapper | `C:\Windows\Logs\Software\Microsoft_Microsoft365Apps*_PSAppDeployToolkit_*.log` |
| Intune agent | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log` |
| Intune agent | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\AgentExecutor.log` |

All of the above are CMTrace format.

Tail the installer log:

```powershell
Get-Content "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log" -Tail 50
```

## Manual test commands

Run from an **elevated 64-bit PowerShell**, inside the `AppInstaller` folder.

```powershell
# Microsoft 365 Apps
.\Files\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Install   -ProductID O365ProPlusRetail -Verbose
.\Files\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Uninstall -ProductID O365ProPlusRetail -Verbose

# Language pack / proofing tools (payload is not in Files\ for these packages)
.\Install-LanguagePacks_v3.0.ps1 -LanguageID "nl-nl" -Mode Install -Verbose
.\Install-ProofingTools_v3.0.ps1 -LanguageID "nl-nl" -Mode Install -Verbose

# PSADT package, from the package root
.\Invoke-AppDeployToolkit.exe -DeploymentType Install   -DeployMode Interactive
.\Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent
```

```powershell
# Current Office state
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' |
    Select-Object ProductReleaseIds, VersionToReport, ClientCulture, Platform, ExecutingScenario

# Installed products and languages in Add/Remove Programs
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' |
    Where-Object PSChildName -match 'O365|Office' |
    ForEach-Object { $_.PSChildName }

# Installed cultures straight from the C2R product keys
$p = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
Get-ChildItem $p | ForEach-Object { Get-ChildItem "$($_.PSPath)\*" | ForEach-Object { $_.Name } }

# Is the engine busy right now?
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').ExecutingScenario
Get-Process setup, OfficeC2RClient -ErrorAction SilentlyContinue
```

## Exit codes you will actually see

| Code | Source | Meaning and what to do |
|---|---|---|
| `0` | any | Success |
| `1` | our installer scripts | Script-level failure — download, signature check, missing XML, or an unhandled error. Read `Microsoft-365-Apps-Setup.log`. |
| `3010` / `1641` | ODT | Reboot required. Map both in the Intune app's Return codes. |
| `1618` | Windows Installer | Another install is in progress. Map to Retry. |
| `60001` | PSADT | Unhandled error inside the wrapper. The PSADT log has the stack trace. If you see this instead of a real ODT code, `-IgnoreExitCodes '*'` is missing from `Start-ADTProcess`. |
| `60008` | PSADT | **The module could not be imported, or a signed toolkit file was edited.** By far the most common packaging mistake — see below. |
| `60012` | PSADT | The user deferred. Only possible if you enabled `-AllowDefer`. Map to Retry or Intune calls it a failure. |
| `69001` | our wrapper | The script named by `$InstallerScriptName` was not found in `Files\`. |
| `17002` / `17004` | ODT | Install cancelled / unknown product ID. Check the Product ID in `Configuration.xml`. |
| `30015` / `30125` | ODT | Download or CDN failure. Usually a proxy or firewall blocking `officecdn.microsoft.com`. |

## "Everything fails with 60008"

Four causes, in order of likelihood:

1. **A file inside `PSAppDeployToolkit\` was hand-edited after being signed** — most often
   `Config\config.psd1` (e.g. pointing `Assets.Banner` at a new filename). PSADT verifies its
   own signed files at startup and refuses to open a session if one doesn't match
   (`ADTDataFileSignatureError`). This can happen silently mid-dialog too: a session that
   fails after the close-apps UI already spawned can leave a stray dialog on screen with
   nothing left to resolve it. `Build-PSADTPackage.ps1` check 8 catches this — run it after
   any edit under `PSAppDeployToolkit\`, and revert to the shipped content (or overwrite an
   `Assets\*.png` file's bytes under the same filename, which isn't signed) instead of
   changing paths in `config.psd1`.
2. **The wrapper was hand-copied from `Frontend\v4` and its module import was never
   corrected.** That file imports from `$PSScriptRoot\..\..\..\PSAppDeployToolkit`, which is
   correct where it normally lives but wrong in a package root. PSADT's own `New-ADTTemplate`
   rewrites it automatically; copying by hand does not. See
   [Packaging](PACKAGING.md#if-you-start-from-the-stock-frontendv4invoke-appdeploytoolkitps1-instead).
   `Build-PSADTPackage.ps1` catches it.
3. **The toolkit module is not inside the package.** Target devices do not have PSADT
   installed; it has to travel with the package.
4. **Your repo path is too long.** Over MAX_PATH the toolkit's own DLLs fail to load. Move
   the repo closer to the drive root.

## "Intune says installed but not detected"

The install returned 0 but the detection script returned non-zero. Almost always a language
that did not install — the detection scripts require both `nl-nl` and `en-us` by default.
Check `Detect-Microsoft-365-Apps.log` and the `Required language` lines in
`Microsoft-365-Apps-Setup.log`. To relax it, set `$RequiredLanguages = @()` at the top of the
detection script.

## "The uninstall does nothing until I reboot"

Fixed in installer v2.5 by `Wait-OfficeEngineIdle`, which blocks until the Click-to-Run
engine has finished its current scenario before starting `setup.exe`. If you still see it,
confirm the log contains `Click-to-Run engine is idle after Ns`. See the
[testing checklist](TESTING.md#4-uninstall-immediately-after-install-without-a-reboot).

## "The uninstall does nothing when an Office app is open" / "the close-apps popup never goes away"

Traced live on 2026-09-05: in both cases the actual symptom was `config.psd1`'s Authenticode
signature being broken by a hand edit, which aborts `Open-ADTSession` before the uninstall
logic ever runs (see cause 1 under "Everything fails with 60008" above). Once the signature is
valid, both the raw installer script (`FORCEAPPSHUTDOWN=True` in `Uninstall.xml`) and the
PSADT close-apps dialog already handle an open Office app correctly — including the dialog
auto-dismissing within seconds when the user closes the app themselves, without needing to
click "Close Programs". Confirmed end-to-end (install → uninstall with Word/Outlook open →
reinstall) as SYSTEM on a live device.

## "It works when I run it manually but fails from Intune"

The Intune Management Extension is a 32-bit process, so a script it launches runs under
WOW64: `$env:ProgramFiles` becomes `Program Files (x86)` and `HKLM\SOFTWARE\Microsoft\Office`
redirects to `WOW6432Node`. Installer v2.5 re-launches itself in 64-bit PowerShell and the
detection scripts read the 64-bit registry view explicitly. Confirm with:

```powershell
Select-String "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log" `
    -Pattern '64-bit process: True'
```
