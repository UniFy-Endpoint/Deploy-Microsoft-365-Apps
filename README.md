# Deploy-Microsoft-365-Apps

#Overview
Deploy Microsoft 365 Apps for Enterprise or Microsoft 365 Apps for Business from Intune as a Win32 App. (Optimized for Autopilot ESP).


#DESCRIPTION
  - Downloads the latest Office Click-to-Run setup.exe version from Microsoft CDN (tiny package).
  - Supports Install and Uninstall modes (default: Install).
  - Uses packaged Configuration.xml for install and packaged Uninstall.xml for removal.
  - Optionally accepts -XMLUrl to fetch a remote XML for either mode.
  - Verifies Microsoft signature on setup.exe.
  - Closes Microsoft 365 apps before Uninstall and removes ProofingTools.
  - Cleans up after completion and returns proper exit codes for Intune.



#Key Features
- Optimized for Autopilot ESP: tiny package, streams bits from Microsoft CDN
- Install and Uninstall support via a single script (Mode parameter)
- Optional -XMLUrl to fetch XML from a remote location
- Uses packaged Configuration.xml for install and Uninstall.xml for removal by default
- Closes all Microsoft 365 apps before Uninstall; removes ProofingTools
- Verifies Microsoft signature on setup.exe
- Cleans up temp files and surfaces proper exit codes for Intune
- Works on AMD64 and ARM64 Windows



#How It Works
- The Script creates a temp work folder (C:\Windows\Temp\OfficeSetup). Downloads setup.exe from the Microsoft URL: https://officecdn.microsoft.com/pr/wsus/setup.exe
  
- Verifies that setup.exe is signed by Microsoft.
  
- Selects the XML option to use (If -XMLUrl is provided, the XML is downloaded to the temp folder):
Install mode: packaged Configuration.xml (unless -XMLUrl is provided)
Uninstall mode: packaged Uninstall.xml (unless -XMLUrl is provided)

- Runs: setup.exe /configure "\configuration.xml"
  
- Cleans up and returns the setup.exe exit code to Intune.

Note: You don’t need to pass the XML path on the Intune command line. The script handles this internally and copies/renames the chosen XML to configuration.xml in the temp folder.



#Create a .intunewin package
- Download the Intune Win32 Content Prep Tool (IntuneWinAppUtil.exe) from GitHub: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/refs/heads/master/IntuneWinAppUtil.exe
  
- Place all the files in the same source folder before you run IntuneWinAppUtil:
  Install-M365-Apps.ps1Install.xml
  Configuration.xml
  Uninstall.xml
  
- Run: IntuneWinAppUtil.exe -c "C:\SourceFolder" -s "Install-M365-Apps.ps1" -o "C:\OutputFolder"
  
- This generates a .intunewin package.



#Create the Win32 App in Intune

- Go to Microsoft Intune Admin Center https://intune.microsoft.com => Apps => Windows => Add.
- Choose App type: Win32 app and Upload the .intunewin file.
- Select the .intunewin you created.
- Use the following install/uninstall commands:

- Install using packaged Configuration.xml
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install

- Uninstall using packaged Uninstall.xml
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall

  
- Install using remote configuration XML
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"

- Uninstall using remote Uninstall XML
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml"



#Detection rules (Custom script)
Use the PowerShell detection script (no version check; supports AMD64 + ARM64).



#Logging and Troubleshooting
Primary log (script): (C:\Windows\Temp\M365-Apps-Setup.log)
Office Click-to-Run logs: Typically under (C:\ProgramData\Microsoft\Office\ClickToRun\Log\)
Common causes of failure: Network egress blocked to officecdn.microsoft.com (XML URL unreachable or invalid)
