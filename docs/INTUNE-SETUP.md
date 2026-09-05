# Intune app settings

Field-by-field settings for creating each app in Intune, and what the wrapper actually runs
underneath the one command you type.

## Creating the app in Intune

1. Go to **Microsoft Intune admin center → Apps → Windows → Add**.
2. Choose app type **Windows app (Win32)** and upload the `.intunewin` file from
   `<App>\Package\`.
3. Fill in Name, Description, Publisher on the **App information** tab. The **Logo** field
   is a manual upload here — the package itself doesn't supply an icon
   (see the `SupportFiles\` note in [Packaging](PACKAGING.md#folder-layout)).
4. On **Program**, enter the install and uninstall commands below, and set
   **Install behavior** to `System`.
5. On **Requirements**, set the OS architecture and minimum OS your estate needs.
6. On **Detection rules**, choose **Use a custom detection script** and upload the matching
   script from that app's `Detection\` folder. Leave *Run script as 32-bit* on **No**.
7. On **Return codes**, add the codes listed below.
8. Assign the app, then verify with the [testing checklist](TESTING.md) before broad
   deployment.

## Microsoft 365 Apps (Business / Enterprise)

PSADT package:

```text
Install   : Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Auto
Uninstall : Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent
```

Win32 script package (Business) — only after converting back per
[Packaging](PACKAGING.md#converting-a-psadt-package-back-to-a-win32-script-package):

```text
Install   : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Install -ProductID O365BusinessRetail
Uninstall : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps_v2.5.ps1 -Mode Uninstall -ProductID O365BusinessRetail
```

Win32 script package (Enterprise): same, with `-ProductID O365ProPlusRetail`.

Overriding the product at deployment time (one package, both editions):

```text
Invoke-AppDeployToolkit.exe -DeploymentType Install   -DeployMode Auto   -ProductID O365ProPlusRetail
Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent -ProductID O365ProPlusRetail
```

### What the wrapper actually runs inside the package

The two Intune commands above are only the entry point. Inside the package,
`Invoke-AppDeployToolkit.ps1` builds and runs the command line below via `Start-ADTProcess`
(function `Invoke-M365InstallerScript`). **You do not type these anywhere** — they are listed
so you can match them against the `Launching [...]` line in the PSADT log.

**Business** — `AppInstaller\Invoke-AppDeployToolkit.ps1`, `$ProductID = 'O365BusinessRetail'`:

| Deployment type | Command the wrapper runs |
|---|---|
| Install | `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<package>\Files\Install-Microsoft-365-Apps_v2.5.ps1" -Mode Install -ProductID O365BusinessRetail` |
| Uninstall | `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<package>\Files\Install-Microsoft-365-Apps_v2.5.ps1" -Mode Uninstall -ProductID O365BusinessRetail` |
| Repair | Same as Install. Re-applying `configuration.xml` with `/configure` is how ODT repairs a Click-to-Run installation. |

**Enterprise** — identical, with `-ProductID O365ProPlusRetail`.

Where the pieces come from:

| Piece | Source |
|---|---|
| `powershell.exe` path | `Get-M365PowerShellPath` — resolves `SysNative` when the toolkit is hosted in a 32-bit process, `System32` otherwise |
| `<package>\Files` | `$adtSession.DirFiles` |
| script name | `$InstallerScriptName` at the top of the wrapper |
| `-Mode` | `Install` / `Uninstall`, from `-DeploymentType` |
| `-ProductID` | the wrapper's `-ProductID` parameter, defaulted per package |

`Start-ADTProcess` is called with `-CreateNoWindow -IgnoreExitCodes '*' -PassThru`, so the
real exit code from `setup.exe` reaches Intune instead of the generic toolkit failure 60001.

To confirm on a device, look for this line in the PSADT log:

```powershell
Select-String "C:\Windows\Logs\Software\Microsoft_Microsoft365Apps*.log" -Pattern 'Launching \['
```

## Language packs / proofing tools

```text
Install   : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-LanguagePacks_v3.0.ps1 -LanguageID "nl-nl" -Mode Install
Uninstall : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-LanguagePacks_v3.0.ps1 -LanguageID "nl-nl" -Mode Uninstall

Install   : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-ProofingTools_v3.0.ps1 -LanguageID "nl-nl" -Mode Install
Uninstall : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-ProofingTools_v3.0.ps1 -LanguageID "nl-nl" -Mode Uninstall
```

If these are later wrapped in PSADT, the Intune commands become the same two lines as above,
with the wrapper passing `-LanguageID` instead of `-ProductID`.

## Common settings

| Setting | Value |
|---|---|
| Install behavior | System |
| Device restart | No specific action |
| Install time | raise 60 → **120** minutes |
| Detection | custom script, Run as 32-bit = **No**, Signature check = **No** |

Detection script per app:

| App | Script |
|---|---|
| Business | `Detect-Microsoft-365-Apps_v2.0.ps1` |
| Enterprise | `Detect-Microsoft-365-Apps_v2.0.ps1` |
| LanguagePacks | `Detect-LanguagePacks_v3.0.ps1` |
| ProofingTools | `Detect-ProofingTools_v3.0.ps1` |

Return codes: `0` and `1707` Success, `3010` Soft reboot, `1641` Hard reboot, `1618` Retry.
Add `60012` = Retry **only** if you enable `-AllowDefer` in the wrapper.
