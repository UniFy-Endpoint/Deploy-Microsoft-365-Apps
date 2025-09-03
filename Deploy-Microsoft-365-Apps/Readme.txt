
It is not required to package the script and the .xml configuration files into an .intunewin file each time before using the package. The file located in the (Package) folder has already been packaged and can be used directly as a Win32 App to configure and install Microsoft 365 Apps. 

A new .intunewin file should only be created if modifications are made to the .xml files or if a new configuration file is required.

- Download the Intune Win32 Content Prep Tool (IntuneWinAppUtil.exe) from GitHub: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/refs/heads/master/IntuneWinAppUtil.exe

To create a package and call it from Intune. Place all the files in the same source folder before you run IntuneWinAppUtil:

- Install-M365-Apps.ps1Install.xml
- Configuration.xml
- Uninstall.xml

Create the .intunewin from that folder.

- Run: IntuneWinAppUtil.exe -c "C:\Source\AppInstaller" -s "Install-M365-Apps.ps1" -o "C:\Destination\Package"

- This generates a .intunewin package.

Display name: Microsoft 365 Apps
Publisher: Microsoft Corporation
Homepage: https://www.office.com/


Install using packaged Configuration.xml: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install

Uninstall using packaged Uninstall.xml: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall


Install using remote configuration XML: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"

Uninstall using remote Uninstall XML: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml"


- Use the Detect-M365-App.ps1 as detection Script.

