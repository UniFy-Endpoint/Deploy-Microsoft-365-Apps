<#
.SYNOPSIS
  Install/Uninstall Microsoft 365 Apps as a Win32 Apps. (Optimized for Autopilot ESP)

.DESCRIPTION
  - Downloads the latest Office Click-to-Run setup.exe from Microsoft CDN (tiny package).
  - Supports Install and Uninstall modes (default: Install).
  - Uses packaged Configuration.xml for install and packaged Uninstall.xml for removal.
  - Optionally accepts -XMLUrl to fetch a remote XML for either mode.
  - Verifies Microsoft signature on setup.exe.
  - Closes Microsoft 365 apps before Uninstall and removes ProofingTools.
  - Cleans up after completion and returns proper exit codes for Intune.

.EXAMPLES
  Install using packaged Configuration.xml
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Install

  Install using remote configuration XML
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -XMLUrl "https://yourdomain.com/office/Configuration.xml"



  Uninstall using packaged Uninstall.xml
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall

  Uninstall using remote Uninstall XML
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-M365-Apps.ps1 -Mode Uninstall -XMLUrl "https://yourdomain.com/office/Uninstall.xml"

.NOTES
  Version: 2.0.0
  Changes:
    - Add Mode (Install/Uninstall)
    - Close M365 apps before Uninstall
    - Remove ProofingTools via Uninstall.xml
    - Keep -XMLUrl supported for both modes
    - Minor robustness + TLS 1.2 default
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [ValidateSet('Install','Uninstall')]
  [string]$Mode = 'Install',

  [Parameter(Mandatory = $false)]
  [string]$XMLUrl,

  [Parameter(Mandatory = $false)]
  [int]$AppCloseTimeoutSeconds = 20
)

# ------------------------- PARAMETERS -------------------------
$LogFileName = "M365-Apps-Setup.log"

function Write-LogEntry {
  param(
    [Parameter(Mandatory=$true)][string]$Value,
    [Parameter(Mandatory=$true)][ValidateSet('1','2','3')][string]$Severity,
    [Parameter(Mandatory=$false)][string]$FileName = $LogFileName
  )
  try {
    $LogFilePath = Join-Path -Path $env:SystemRoot -ChildPath ("Temp\" + $FileName)
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $Date = (Get-Date -Format "MM-dd-yyyy")
    $Context = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
    Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    if ($Severity -eq 1) { Write-Verbose $Value } elseif ($Severity -eq 3) { Write-Warning $Value }
  } catch {
    Write-Warning "Unable to append log entry to $LogFileName. Error: $($_.Exception.Message)"
  }
}

function Set-Tls12 {
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-LogEntry -Value "TLS set to TLS 1.2" -Severity 1
  } catch {
    Write-LogEntry -Value "Failed to set TLS 1.2: $($_.Exception.Message)" -Severity 2
  }
}

function Start-DownloadFile {
  param(
    [Parameter(Mandatory=$true)][string]$URL,
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$Name
  )
  if (-not (Test-Path -Path $Path)) {
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
  }
  $dest = Join-Path -Path $Path -ChildPath $Name
  try {
    # Prefer Invoke-WebRequest, fallback to WebClient if needed
    try {
      Invoke-WebRequest -Uri $URL -OutFile $dest -UseBasicParsing -ErrorAction Stop
    } catch {
      $wc = New-Object System.Net.WebClient
      $wc.DownloadFile($URL, $dest)
      $wc.Dispose()
    }
    return $dest
  } catch {
    throw "Download failed from $URL to $dest. Error: $($_.Exception.Message)"
  }
}

function Invoke-FileCertVerification {
  param([Parameter(Mandatory=$true)][string]$FilePath)
  $sig = Get-AuthenticodeSignature -FilePath $FilePath
  $Cert = $sig.SignerCertificate
  $CertStatus = $sig.Status
  if ($Cert) {
    if ($Cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid") {
      $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
      $null = $chain.Build($Cert)
      $RootCert = $chain.ChainElements | ForEach-Object { $_.Certificate } | Where-Object { $_.Subject -match "CN=Microsoft Root" }
      if ($RootCert) {
        $TrustedRoot = Get-ChildItem "Cert:\LocalMachine\Root" -Recurse | Where-Object { $_.Thumbprint -eq $RootCert.Thumbprint }
        if ($TrustedRoot) {
          Write-LogEntry -Value "Verified setup file signed by: $($Cert.Issuer)" -Severity 1
          return $true
        } else {
          Write-LogEntry -Value "No trust found to root cert - aborting" -Severity 2
          return $false
        }
      } else {
        Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2
        return $false
      }
    } else {
      Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2
      return $false
    }
  } else {
    Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
    return $false
  }
}

function Stop-M365Apps {
  param([int]$TimeoutSeconds = 20)
  # Attempt graceful close, then force if still running
  $procNames = @(
    'WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','ONENOTEM',
    'MSACCESS','MSPUB','VISIO','PROJECT','LYNC','C2RClient'
  )
  $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $procNames -contains $_.Name }
  if (-not $procs) {
    Write-LogEntry -Value "No active Microsoft 365 app processes detected" -Severity 1
    return
  }
  Write-LogEntry -Value ("Attempting to close Microsoft 365 apps: " + ($procs.Name | Sort-Object -Unique -join ', ')) -Severity 1
  foreach ($p in $procs) {
    try { $null = $p.CloseMainWindow() } catch {}
  }
  Start-Sleep -Seconds ([Math]::Min([Math]::Max($TimeoutSeconds, 5), 60))

  # Force kill remaining
  $stillRunning = Get-Process -ErrorAction SilentlyContinue | Where-Object { $procNames -contains $_.Name }
  foreach ($p in $stillRunning) {
    try {
      Stop-Process -Id $p.Id -Force -ErrorAction Stop
      Write-LogEntry -Value "Force-closed $($p.Name) (PID $($p.Id))" -Severity 2
    } catch {
      Write-LogEntry -Value "Failed to close $($p.Name): $($_.Exception.Message)" -Severity 2
    }
  }
}

# ------------------------- Microsoft 365 Apps Installation -------------------------
Set-Tls12
Write-LogEntry -Value "Mode: $Mode. Starting M365 Apps setup process" -Severity 1

# Prep temp working folder
$workRoot = Join-Path $env:SystemRoot "Temp\OfficeSetup"
if (Test-Path $workRoot) {
  Remove-Item -Path $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}
$null = New-Item -ItemType Directory -Path $workRoot -Force

# Download latest Office Click-to-Run setup.exe
$setupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
Write-LogEntry -Value "Downloading Office setup from evergreen URL" -Severity 1
try {
  $setupPath = Start-DownloadFile -URL $setupUrl -Path $workRoot -Name "setup.exe"
  if (-not (Test-Path $setupPath)) { throw "setup.exe not found after download" }
  Write-LogEntry -Value "setup.exe downloaded to $setupPath" -Severity 1
} catch {
  Write-LogEntry -Value "Error downloading setup.exe: $($_.Exception.Message)" -Severity 3
  exit 1
}

# Verify signature
if (-not (Invoke-FileCertVerification -FilePath $setupPath)) {
  Write-LogEntry -Value "Signature verification failed for setup.exe" -Severity 3
  exit 1
}

# Prepare configuration XML
$configPath = Join-Path $workRoot "configuration.xml"
try {
  if ($XMLUrl) {
    Write-LogEntry -Value "Fetching configuration XML from URL: $XMLUrl" -Severity 1
    $null = Start-DownloadFile -URL $XMLUrl -Path $workRoot -Name "configuration.xml"
  } else {
    if ($Mode -eq 'Install') {
      $src = Join-Path $PSScriptRoot "Configuration.xml"  # packaged with script
    } else {
      $src = Join-Path $PSScriptRoot "Uninstall.xml"      # packaged with script
    }
    if (-not (Test-Path $src)) { throw "Required XML not found in package: $src" }
    Copy-Item $src $configPath -Force -ErrorAction Stop
    Write-LogEntry -Value "Using packaged XML: $(Split-Path $src -Leaf)" -Severity 1
  }
} catch {
  Write-LogEntry -Value "Failed to prepare configuration XML: $($_.Exception.Message)" -Severity 3
  exit 1
}

# Uninstall pre-steps: ensure apps are closed
if ($Mode -eq 'Uninstall') {
  Write-LogEntry -Value "Closing Microsoft 365 apps before uninstall" -Severity 1
  Stop-M365Apps -TimeoutSeconds $AppCloseTimeoutSeconds
}

# Run setup.exe /configure
try {
  $ver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($setupPath).FileVersion
  Write-LogEntry -Value "Office C2R setup version $ver" -Severity 1
} catch {}

Write-LogEntry -Value "Starting setup.exe /configure configuration.xml" -Severity 1
try {
  $proc = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait -PassThru -ErrorAction Stop
  $exitCode = $proc.ExitCode
  Write-LogEntry -Value "setup.exe completed with exit code: $exitCode" -Severity 1
} catch {
  Write-LogEntry -Value "Error running setup.exe: $($_.Exception.Message)" -Severity 3
  $exitCode = 1
}

# Cleanup
try {
  if (Test-Path $workRoot) {
    Remove-Item -Path $workRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
} catch {}

if ($exitCode -ne 0) {
  Write-LogEntry -Value "M365 Apps $Mode failed" -Severity 3
  exit $exitCode
}

Write-LogEntry -Value "M365 Apps $Mode completed successfully" -Severity 1
exit 0