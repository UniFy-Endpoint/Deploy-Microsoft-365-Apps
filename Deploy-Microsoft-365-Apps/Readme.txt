- Download the Intune Win32 Content Prep Tool (IntuneWinAppUtil.exe) from GitHub: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/refs/heads/master/IntuneWinAppUtil.exe

To create a package and call it from Intune. Place all the files in the same source folder before you run IntuneWinAppUtil:

- Install-M365-Apps.ps1Install.xml
- Configuration.xml
- Uninstall.xml

Create the .intunewin from that folder.

- Run: IntuneWinAppUtil.exe -c "C:\Intune\Apps\Install-M365-Apps\AppInstaller" -s "Install-M365-Apps.ps1" -o "C:\Intune\Apps\Install-M365-Apps\Package"

- This generates a .intunewin package.

Display name: Microsoft 365 Apps for Enterprise
Publisher: Microsoft Corporation
Homepage: https://www.office.com/



Install using packaged Configuration.xml
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install

Install using remote configuration XML
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"



Uninstall using packaged Uninstall.xml
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall

Uninstall using remote Uninstall XML
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml


- Use the Detect-M365-App.ps1 as detection Script.