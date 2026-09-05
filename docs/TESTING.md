# Validation checklist

Run on a test device. Tick every box before releasing to production.

Steps marked **[!]** exist because an earlier version of these scripts got that specific
thing wrong in a way that was invisible until it reached real devices. They are kept as
checks because they are also the failures most likely to appear in a freshly assembled
package. Each one names the original defect so you know what you are looking for — run them
even if you are starting from a clean clone.

## 1. Detection scripts, before installing anything

- [ ] Business detection on a clean device returns exit `1`
      `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Detect-Microsoft-365-Apps_v2.0.ps1; $LASTEXITCODE`
- [ ] It writes **nothing** to stdout when not detected (Intune needs exit 0 **and** stdout to mark an app as detected)
- [ ] Language pack detection returns exit `1`
- [ ] Proofing tools detection returns exit `1`

## 2. Install (Microsoft 365 Apps)

- [ ] Install runs to completion and returns exit `0`
- [ ] Log confirms the 64-bit context (this was the v2.3/v2.4 defect):
      `Select-String "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log" -Pattern '64-bit process: True'`
- [ ] **[!]** Log does **not** contain repeated `Office registry not found yet, waiting...` lines.
      Those mean the script is running under WOW64 and every verification is failing.
- [ ] Log shows `Office installation verified complete. Version: ...`
- [ ] Log shows `ClickToRunSvc service stopped successfully`
- [ ] **[!]** Log lists **both** required languages — `Select-String <log> -Pattern "Required language"`
      Expect `Required language 'nl-nl' is installed` and `'en-us' is installed`
- [ ] Detection now returns exit `0` and prints one line

## 3. Product isolation (this was a real false positive)

- [ ] **[!]** On a device with the **other** edition installed, the detection script returns exit `1`.
      Verified during development: a device with `O365ProPlusRetail` was reported as *detected*
      by the old Business detection script, because v1.0 matched either product ID and accepted
      "any DisplayName like `*Microsoft Office*`".

      ```powershell
      Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' |
          Select-Object ProductReleaseIds, VersionToReport, ClientCulture
      ```
- [ ] Uninstall removes **only** the targeted product; any other Office edition survives

      ```powershell
      Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' |
          Where-Object PSChildName -match 'O365|Office'
      ```

## 4. Uninstall immediately after install, without a reboot

> **[!]** This is the problem that used to require a restart. v2.5 waits for the Click-to-Run
> engine to go idle before running `setup.exe`, instead of firing into a running scenario.
> Confirmed working end-to-end (install → uninstall with an Office app open → reinstall) as
> SYSTEM on a live device — see the PSADT behaviour notes in section 6 below.

- [ ] Install, then run the uninstall straight away without rebooting
- [ ] Log shows `Waiting for the Click-to-Run engine to become idle...`
- [ ] Log shows `Click-to-Run engine is idle after Ns`
- [ ] Log shows `<ProductID> successfully removed`
- [ ] Uninstall returns exit `0`
- [ ] Detection returns exit `1` afterwards
- [ ] Repeat the uninstall on a device where the product is **not** installed:
      expect `not detected - nothing to uninstall` and exit `0`

## 5. Language packs and proofing tools

> **[!]** v2.1 of both installers called a function name that did not exist
> (`Test-LanguagePackInstallation` / `Test-ProofingToolsInstallation`), so **every** install
> reported failure to Intune even though the language pack installed correctly.

- [ ] Language pack install returns exit `0` (not `1`)
- [ ] Log shows `... verified successfully` — **not** `Critical error ... is not recognized as the name of a cmdlet`
- [ ] Language pack detection returns exit `0` afterwards
- [ ] **[!]** Proofing tools detection still returns exit `1` when **only** the language pack is
      installed. The old script reported the proofing tools app as installed in that case.
- [ ] Install the proofing tools, then their detection returns exit `0`
- [ ] Confirm the two are distinguishable in the registry:

      ```powershell
      Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' |
          Where-Object PSChildName -match 'nl-nl'
      ```

      Language pack → `O365ProPlusRetail - nl-nl` · Proofing → `O365ProPlusRetail - nl-nl.proof`

## 6. PSADT behaviour

- [ ] With no Office app running: no dialog appears, install proceeds immediately
- [ ] With Word open and a user logged on: countdown dialog appears listing Word,
      with a **Close Programs** button and **no Postpone** button
- [ ] Clicking **Close Programs** closes Word and the install continues
- [ ] **Closing the app yourself** (not clicking the button) — the dialog detects the
      process is gone and dismisses on its own within a couple of seconds; confirm in the log
      with `The user selected to continue...`. Confirmed on a live device 2026-09-05.
- [ ] At the end of the countdown (if left untouched) Word is closed and the install continues
- [ ] Running as SYSTEM with no user logged on: log shows
      `deployment mode was explicitly set to [Silent]` and no dialog is shown
- [ ] Autopilot / ESP: log shows OOBE or an active ESP, and the deployment runs Silent
- [ ] PSADT exit code equals the installer script exit code, not `60001`
- [ ] `PSAppDeployToolkit\Config\config.psd1` and every `.psd1`/`.psm1`/`.ps1` under
      `PSAppDeployToolkit\` still carry a **valid** Authenticode signature (`Build-PSADTPackage.ps1`
      check 8). A broken signature aborts the session before the dialog can ever resolve —
      this is what "the close-apps popup never goes away" and "uninstall does nothing with an
      app open" both turned out to be, traced live on 2026-09-05.

## 7. Autopilot end-to-end

- [ ] Reset the device to factory settings and run a full Autopilot enrolment
- [ ] Office is present and usable at first user logon
- [ ] Intune reports the app as installed (not "installed but not detected")
- [ ] Word UI language matches the OS language; `nl-nl` and `en-us` are both available
      under **File → Options → Language**
