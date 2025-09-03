# Deploy Microsoft 365 Apps with Intune (Win32 App)

---

## Overview

Deploy **Microsoft 365 Apps for Enterprise** or **Microsoft 365 Apps for Business** from **Intune** as a Win32 App (optimized for Autopilot ESP).  

You do **not** need to package the script and the `.xml` configuration files into an `.intunewin` file every time. The file located in the **Package** folder is already packaged and can be used directly as a Win32 App to configure and install Microsoft 365 Apps.  

A new `.intunewin` file should only be created if:  
- Modifications are made to the `.xml` files, or  
- A new configuration file is required.  

---

## Important Notice

It is strongly recommended to **validate and test** both the script and the configuration in a **controlled test environment** before deploying in production.  

---

## Description

This Script:  
- Downloads the latest Office Click-to-Run `setup.exe` version from the Microsoft CDN (tiny package).  
- Supports **Install** and **Uninstall** modes (default = Install).  
- Uses **Configuration.xml** for install and **Uninstall.xml** for removal (packaged by default).  
- Optionally accepts `-XMLUrl` to fetch a remote XML for either mode.  
- Verifies the Microsoft signature on `setup.exe`.  
- Closes Microsoft 365 apps before uninstalling and removes Proofing Tools.  
- Cleans up after completion and returns proper exit codes for Intune.  

---

## Key Features

- Optimized for **Autopilot ESP** (tiny package, streams bits from Microsoft CDN)  
- A single script supports both Install/Uninstall modes via the `-Mode` parameter  
- Optional `-XMLUrl` to fetch configuration XML files remotely  
- Defaults to packaged `Configuration.xml` (Install) and `Uninstall.xml` (Uninstall)  
- Verifies Microsoft’s digital signature on `setup.exe`  
- Cleans up temporary files and ensures proper exit codes surface in Intune  
- Works on **AMD64** and **ARM64** Windows devices  

---

## How It Works

1. The script creates a temporary folder: C:\Windows\Temp\OfficeSetup
2. Downloads `setup.exe` from: https://officecdn.microsoft.com/pr/wsus/setup.exe
3. Verifies that `setup.exe` is signed by Microsoft.  
4. Selects an XML option:  
- **Install mode**: uses packaged `Configuration.xml` (or downloads via `-XMLUrl`)  
- **Uninstall mode**: uses packaged `Uninstall.xml` (or downloads via `-XMLUrl`)  
5. Cleans up and returns the `setup.exe` exit code to Intune.  

**Note:** You do not need to pass the XML path in the Intune command line. The script automatically copies/renames the chosen XML to `configuration.xml` in the temp folder.  

---

## Create a .intunewin Package (Optional)

1. Download the [Intune Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/refs/heads/master/IntuneWinAppUtil.exe).  
2. Place all required files in the same source folder:  
- Install-M365-Apps.ps1
- Configuration.xml
- Uninstall.xml

3. Run the following powershell command to generate a .intunewin package:  
.\IntuneWinAppUtil.exe -c "C:\SourceFolder" -s "Install-M365-Apps.ps1" -o "C:\OutputFolder"

---
## Create a Win32 App in Intune

1. Go to Microsoft Intune Admin Center → Apps → Windows → Add.
2. Choose App type: Win32 app and upload the .intunewin file.
3. Use the following install/uninstall commands:

- Install (using packaged Configuration.xml): powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install
- Uninstall (using packaged Uninstall.xml): powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall

---

- Install (using remote Configuration.xml): powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"
- Uninstall (using remote Uninstall.xml): powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml"

---
## Detection Rules

- Use the PowerShell detection script (no version check required).
- Supports AMD64 and ARM64 architectures.

---
## Logging and Troubleshooting

- Script log: C:\Windows\Temp\M365-Apps-Setup.log
- Office Click-to-Run logs: C:\ProgramData\Microsoft\Office\ClickToRun\Log
- Common causes of failure: Network egress blocked to officecdn.microsoft.com or the provided XML URL is unreachable or invalid.
