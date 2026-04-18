<#
.SYNOPSIS
  Script to install additional Languagepacks as a Win32 App 

.DESCRIPTION
    Script to install additional languagepacks as a Win32 App by downloading the latest Office Deployment Toolkit
    Running Setup.exe from downloaded files with provided install.xml and uninstall.xml files.
    NOW INCLUDES: Proper wait logic to ensure installation completes before script exits

.PARAMETER LanguageID
    Set the language ID in the correct formatting (like nl-nl or en-us)
.PARAMETER Mode 
    Supported modes are Install or Uninstall 

.EXAMPLE 
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-LanguagePacks.ps1 -LanguageID "nl-nl" -Mode Install
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-LanguagePacks.ps1 -LanguageID "nl-nl" -Mode Uninstall

.NOTES
  Version:    2.1 (Fixed loop issue + Uninstall verification)
  Author:    UniFy-Endpoint
  
#>

#region parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [string]$LanguageID,

    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxWaitMinutes = 30
)
#endregion parameters

#Region Functions
function Write-LogEntry {
    param (
    [parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Value,
    [parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("1", "2", "3")]
    [string]$Severity,
    [parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$FileName = $LogFileName
    )
    $LogFilePath = Join-Path -Path $env:SystemRoot -ChildPath $("Temp\$FileName")
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $Date = (Get-Date -Format "MM-dd-yyyy")
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
    
    try {
    Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    if ($Severity -eq 1) {
    Write-Verbose -Message $Value
    } elseif ($Severity -eq 3) {
    Write-Warning -Message $Value
    }
    } catch [System.Exception] {
    Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

function Start-DownloadFile {
    param(
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$URL,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name
    )
    Begin {
    # Use modern Invoke-WebRequest instead of deprecated WebClient
    }
    Process {
    if (-not(Test-Path -Path $Path)) {
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    $DestinationPath = Join-Path -Path $Path -ChildPath $Name
    Invoke-WebRequest -Uri $URL -OutFile $DestinationPath -UseBasicParsing
    }
    End {
    # No cleanup needed for Invoke-WebRequest
    }
}

function Invoke-XMLUpdate {
    param(
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$LanguageID,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Filename,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode 
    )
    if ($Mode -eq "Install"){
    $xmlDoc = [System.Xml.XmlDocument](Get-Content $FileName)
    $xmlDoc.Configuration.Add.Product.Language.ID = $LanguageID
    $xmlDoc.Save($FileName); 
    }
    else {
    $xmlDoc = [System.Xml.XmlDocument](Get-Content $FileName)
    $xmlDoc.Configuration.Remove.Product.Language.ID = $LanguageID
    $xmlDoc.Save($FileName);
    }
}

function Invoke-FileCertVerification {
    param(
    [parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FilePath
    )
    $Cert = (Get-AuthenticodeSignature -FilePath $FilePath).SignerCertificate
    $CertStatus = (Get-AuthenticodeSignature -FilePath $FilePath).Status
    if ($Cert){
    if ($cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid"){
    $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
    $chain.Build($cert) | Out-Null
    $RootCert = $chain.ChainElements | ForEach-Object {$_.Certificate}| Where-Object {$PSItem.Subject -match "CN=Microsoft Root"}
    if (-not [string ]::IsNullOrEmpty($RootCert)){
    $TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $PSItem.Thumbprint -eq $RootCert.Thumbprint}
    if (-not [string]::IsNullOrEmpty($TrustedRoot)){
    Write-LogEntry -Value "Verified setupfile signed by : $($Cert.Issuer)" -Severity 1
    Return $True
    }
    else {
    Write-LogEntry -Value  "No trust found to root cert - aborting" -Severity 2
    Return $False
    }
    }
    else {
    Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2 
    Return $False
    }
    }
    else {
    Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2 
    Return $False
    }  
    }
    else {
    Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
    Return $False
    }
}

function Test-LanguagePackStatus {
    param(
    [parameter(Mandatory=$true)]
    [string]$LanguageID,
    [parameter(Mandatory=$true)]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode,
    [parameter(Mandatory=$false)]
    [int]$MaxWaitMinutes = 30
    )
    
    Write-LogEntry -Value "Starting Language Pack $Mode verification for language: $LanguageID (Max wait: $MaxWaitMinutes minutes)" -Severity 1
    
    $StartTime = Get-Date
    $TimeoutTime = $StartTime.AddMinutes($MaxWaitMinutes)
    $CheckIntervalSeconds = 15
    $StableCheckCount = 0
    $RequiredStableChecks = 2  # Require 2 consecutive stable checks
    
    while ((Get-Date) -lt $TimeoutTime) {
    Write-LogEntry -Value "Checking Language Pack $Mode status for $LanguageID..." -Severity 1
    
    # Check for active setup processes (NOT OfficeClickToRun service - it always runs)
    # Only check for setup.exe and OfficeC2RClient which indicate active installation
    $ActiveSetupProcesses = Get-Process -Name "setup", "OfficeC2RClient" -ErrorAction SilentlyContinue
    
    if ($ActiveSetupProcesses) {
    Write-LogEntry -Value "Office setup still in progress (processes: $($ActiveSetupProcesses.Name -join ', ')). Waiting..." -Severity 1
    $StableCheckCount = 0
    Start-Sleep -Seconds $CheckIntervalSeconds
    continue
    }
    
    # Check registry for installed language packs
    $LanguagePackInstalled = $false
    
    # Check ClickToRun Configuration
    $C2RConfigPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $C2RConfigPath) {
    $C2RConfig = Get-ItemProperty -Path $C2RConfigPath -ErrorAction SilentlyContinue
    if ($C2RConfig.InstalledLanguages -match $LanguageID) {
    $LanguagePackInstalled = $true
    Write-LogEntry -Value "Found $LanguageID in ClickToRun InstalledLanguages" -Severity 1
    }
    
    # Also check ProductReleaseIds for LanguagePack
    $ProductReleaseIds = $C2RConfig.ProductReleaseIds
    if ($ProductReleaseIds -match "LanguagePack") {
    Write-LogEntry -Value "LanguagePack found in ProductReleaseIds: $ProductReleaseIds" -Severity 1
    }
    }
    
    # Check Win32 Uninstall registry for language pack entries
    $UninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    foreach ($UninstallPath in $UninstallPaths) {
    if (Test-Path $UninstallPath) {
    # Look for language pack entries (not .proof entries)
    $LanguageEntries = Get-ChildItem -Path $UninstallPath -ErrorAction SilentlyContinue | 
    Get-ItemProperty -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName -match $LanguageID -and $_.DisplayName -notmatch "\.proof" -and $_.DisplayName -match "Microsoft (365|Office)" }
    
    if ($LanguageEntries) {
    $LanguagePackInstalled = $true
    Write-LogEntry -Value "Found language pack in uninstall registry: $($LanguageEntries.DisplayName -join ', ')" -Severity 1
    }
    }
    }
    
    # Determine success based on mode
    if ($Mode -eq "Install") {
    if ($LanguagePackInstalled) {
    $StableCheckCount++
    Write-LogEntry -Value "Language Pack for $LanguageID detected (stable check $StableCheckCount of $RequiredStableChecks)" -Severity 1
    
    if ($StableCheckCount -ge $RequiredStableChecks) {
    Write-LogEntry -Value "Language Pack $LanguageID installation verified successfully!" -Severity 1
    return $true
    }
    }
    else {
    $StableCheckCount = 0
    }
    }
    elseif ($Mode -eq "Uninstall") {
    if (-not $LanguagePackInstalled) {
    $StableCheckCount++
    Write-LogEntry -Value "Language Pack for $LanguageID not found (stable check $StableCheckCount of $RequiredStableChecks)" -Severity 1
    
    if ($StableCheckCount -ge $RequiredStableChecks) {
    Write-LogEntry -Value "Language Pack $LanguageID uninstallation verified successfully!" -Severity 1
    return $true
    }
    }
    else {
    $StableCheckCount = 0
    Write-LogEntry -Value "Language Pack for $LanguageID still present. Waiting for removal..." -Severity 1
    }
    }
    
    $ElapsedMinutes = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 1)
    Write-LogEntry -Value "Elapsed time: $ElapsedMinutes minutes. Waiting $CheckIntervalSeconds seconds..." -Severity 1
    Start-Sleep -Seconds $CheckIntervalSeconds
    }
    
    Write-LogEntry -Value "Timeout reached after $MaxWaitMinutes minutes. Language Pack $Mode may not have completed." -Severity 3
    return $false
}

function Stop-ClickToRunService {
    try {
    $ClickToRunService = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
    if ($ClickToRunService -and $ClickToRunService.Status -eq "Running") {
    Write-LogEntry -Value "Stopping Microsoft Office Click-to-Run service" -Severity 1
    Stop-Service -Name "ClickToRunSvc" -Force -ErrorAction Stop
    Write-LogEntry -Value "Microsoft Office Click-to-Run service stopped successfully" -Severity 1
    } else {
    Write-LogEntry -Value "Microsoft Office Click-to-Run service is not running or not found" -Severity 1
    }
    }
    catch {
    Write-LogEntry -Value "Failed to stop Microsoft Office Click-to-Run service: $($_.Exception.Message)" -Severity 2
    }
}
#Endregion Functions

#Region Initialisations
$LogFileName = "M365LanguagePackSetup.log"
$ExitCode = 0
#Endregion Initialisations

switch -Wildcard ($Mode) { 
    {($PSItem -match "Install")}{
    $FileName = "install.xml"
    }
    {($PSItem -match "Uninstall")}{
    $FileName = "uninstall.xml"
    }
}

# Initate Install
Write-LogEntry -Value "====" -Severity 1
Write-LogEntry -Value "Initiating LanguagePack $($LanguageID) $($Mode) process (v2.1 - Fixed)" -Severity 1
Write-LogEntry -Value "Max wait time: $MaxWaitMinutes minutes" -Severity 1
Write-LogEntry -Value "====" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    $OfficeProcesses = Get-Process -Name "OfficeClickToRun", "setup", "OfficeC2RClient" -ErrorAction SilentlyContinue
    if (-not $OfficeProcesses) {
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
    Write-LogEntry -Value "Cleaned up previous setup folder" -Severity 1
    }
    else {
    Write-LogEntry -Value "Office processes detected, skipping cleanup" -Severity 1
    }
}

$SetupFolder = (New-Item -ItemType "directory" -Path "$($env:SystemRoot)\Temp" -Name OfficeSetup -Force).FullName

try{
    # Download latest Office Deployment Toolkit
    $SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    
    try{
    # Start install preparations
    $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
    if (-Not (Test-Path $SetupFilePath)) {
    Throw "Error: Setup file not found"
    }
    Write-LogEntry -Value "Setup file found at $($SetupFilePath)" -Severity 1
    
    try{
    # Prepare language pack installation or removal
    $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion 
    Write-LogEntry -Value "Office C2R Setup is running version $OfficeCR2Version" -Severity 1
    
    if (Invoke-FileCertVerification -FilePath $SetupFilePath){
    Invoke-XMLUpdate -LanguageID $LanguageID -Filename "$($PSScriptRoot)\$($Filename)" -Mode $Mode
    Copy-Item "$($PSScriptRoot)\$($Filename)" $SetupFolder -Force -ErrorAction Stop
    Write-LogEntry -Value "LanguagePack $($LanguageID) configuration file copied" -Severity 1    
    
    Try{
    # Running office installer
    Write-LogEntry -Value "Starting LanguagePack $($LanguageID) $($Mode) with Win32App method" -Severity 1
    $OfficeInstall = Start-Process $SetupFilePath -ArgumentList "/configure `"$($SetupFolder)\$($Filename)`"" -NoNewWindow -Wait -PassThru -ErrorAction Stop
    
    Write-LogEntry -Value "Setup.exe completed with exit code: $($OfficeInstall.ExitCode)" -Severity 1
    
    if ($OfficeInstall.ExitCode -eq 0) {
    if ($Mode -eq "Install") {
    Write-LogEntry -Value "Setup.exe initiated successfully. Now waiting for actual Language Pack installation to complete..." -Severity 1
    
    # Wait for actual installation to complete
    $InstallSuccess = Test-LanguagePackInstallation -LanguageID $LanguageID -MaxWaitMinutes $MaxWaitMinutes
    
    if ($InstallSuccess) {
    Write-LogEntry -Value "LanguagePack $($LanguageID) installation completed and verified successfully!" -Severity 1
    Stop-ClickToRunService
    }
    else {
    Write-LogEntry -Value "LanguagePack $($LanguageID) installation verification failed or timed out" -Severity 3
    $ExitCode = 1
    }
    }
    else {
    Write-LogEntry -Value "LanguagePack $($LanguageID) $($Mode) completed successfully" -Severity 1
    }
    }
    else {
    Write-LogEntry -Value "LanguagePack $($LanguageID) $($Mode) failed with exit code: $($OfficeInstall.ExitCode)" -Severity 3
    $ExitCode = 1
    }
    }
    catch [System.Exception]{
    Write-LogEntry -Value  "Error running the LanguagePack $($LanguageID) $($Mode). Errormessage: $($_.Exception.Message)" -Severity 3
    $ExitCode = 1
    }
    }
    else {
    Write-LogEntry -Value "Error: Unable to verify setup file signature" -Severity 3
    $ExitCode = 1
    }
    }
    catch [System.Exception]{
    Write-LogEntry -Value  "Error preparing LanguagePack $($LanguageID) $($Mode). Errormessage: $($_.Exception.Message)" -Severity 3
    $ExitCode = 1
    }
    }
    catch [System.Exception]{
    Write-LogEntry -Value  "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3
    $ExitCode = 1
    }
}
catch [System.Exception]{
    Write-LogEntry -Value  "Error downloading setup.exe from evergreen url. Errormessage: $($_.Exception.Message)" -Severity 3
    $ExitCode = 1
}

# Cleanup 
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    $OfficeProcesses = Get-Process -Name "OfficeClickToRun", "setup", "OfficeC2RClient" -ErrorAction SilentlyContinue
    if (-not $OfficeProcesses) {
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
    Write-LogEntry -Value "Cleanup completed" -Severity 1
    }
    else {
    Write-LogEntry -Value "Office processes still running, skipping cleanup to avoid interruption" -Severity 1
    }
}

Write-LogEntry -Value "====" -Severity 1
Write-LogEntry -Value "LanguagePack $($LanguageID) $($Mode) completed with exit code: $ExitCode" -Severity 1
Write-LogEntry -Value "====" -Severity 1

exit $ExitCode