# Deploy Microsoft 365 Apps with Intune (Win32 App)

---

## Overview

Deploy **Microsoft 365 Apps for Enterprise** or **Microsoft 365 Apps for Business** from **Intune** as a Win32 App (optimized for Autopilot ESP).

---

## Important Notice

It is recommended to **validate and test** both the script and the configuration in a **controlled test environment** before deploying in production.  

---

## Description

Script to install or uninstall Microsoft 365 Apps as a Win32 App during Autopilot by downloading the latest Office setup exe from Microsoft CDN url and running Setup.exe with provided configuration.xml or uninstall.xml file.

---

## Key Features

- Optimized for **Autopilot ESP** (tiny package, streams bits from Microsoft CDN)  
- A single script supports **Install** and **Uninstall** modes via the `-Mode` parameter  
- Verifies Microsoft’s digital signature on `setup.exe`  
- Supports both **O365ProPlusRetail** and **O365BusinessRetail** via parameter
- Detects and works **AMD64** **ARM64** Windows system architecture
- Uses **Configuration.xml** for install and **Uninstall.xml**
- Cleans up temporary files and ensures proper exit codes surface in Intune

---

## How It Works

1. The script creates a temporary folder: `C:\Windows\Temp\OfficeSetup`
2. Downloads `setup.exe` from: `https://officecdn.microsoft.com/pr/wsus/setup.exe`
3. Verifies that `setup.exe` is signed by Microsoft.  
4. Start the installation on System Contex.
5. Cleans up temporary files and ensures proper exit codes after installation cpleted.  
  
**Note:** You do not need to pass the XML path in the Intune command line. The script automatically copies/renames the chosen XML to `configuration.xml` to the temp folder `C:\Windows\Temp\OfficeSetup`

---

## Create a .intunewin Package

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

## Install Microsoft 365 Apps for Enterprise

- Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install -ProductID O365ProPlusRetail
- Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall -ProductID O365ProPlusRetail


 ## Install Microsoft 365 Apps for Business
 
- Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install -ProductID O365BusinessRetail
- Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall -ProductID O365BusinessRetail


## Installation Method Without -ProductID (uses XML default)
- Install command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Install
- Uninstall command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-Microsoft-365-Apps.ps1 -Mode Uninstall

---

## Detection Rules

- Use the PowerShell detection script **Detect-Microsoft-365-Apps.ps1**
- Supports AMD64 and ARM64 architectures.

---

## Logging and Troubleshooting

- Script logs: Microsoft\IntuneManagementExtension\Logs\Microsoft-365-Apps-Setup.log"

