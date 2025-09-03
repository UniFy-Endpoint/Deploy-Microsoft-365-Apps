# Deploy-Microsoft-365-Apps
Deploy Microsoft 365 Apps for Enterprise or Microsoft 365 Apps for Business as a Win32 App. (Optimized for Autopilot ESP)

.DESCRIPTION
  - Downloads the latest Office Click-to-Run setup.exe from Microsoft CDN (tiny package).
  - Supports Install and Uninstall modes (default: Install).
  - Uses packaged Configuration.xml for install and packaged Uninstall.xml for removal.
  - Optionally accepts -XMLUrl to fetch a remote XML for either mode.
  - Verifies Microsoft signature on setup.exe.
  - Closes Microsoft 365 apps before Uninstall and removes ProofingTools.
  - Cleans up after completion and returns proper exit codes for Intune.

.EXAMPLES
- Install using packaged Configuration.xml
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install

- Uninstall using packaged Uninstall.xml
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall

  
- Install using remote configuration XML
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"

- Uninstall using remote Uninstall XML
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml"

