<#
.SYNOPSIS
    Detection Script for Microsoft 365 Proofing Tools Language Pack

.DESCRIPTION
    This script is meant to be used with Intune (Win32 app deployment).
    It checks whether the specified Microsoft 365 Proofing Tools language pack is installed.
    Detection is based on Office Click-to-Run configuration and presence of language pack registry keys.

.PARAMETER LanguageID
    Define the Proofing Tools Language ID/Tag (e.g. "en-us", "nl-nl", "de-de")
    
.NOTES
    Version: 2.0 (Fixed)
    Author: UniFy-Endpoint
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$LanguageID = "nl-nl"
)

function Write-DetectionLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    Write-Output "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Test-ProofingToolsInstalled {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Lang
    )

    $Installed = $false
    Write-DetectionLog "Starting Proofing Tools detection for language: $Lang"

    # Method 1: Check Office Click-to-Run Configuration registry
    Write-DetectionLog "Method 1: Checking Office Click-to-Run Configuration..."
    $OfficeConfigPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )

    foreach ($configPath in $OfficeConfigPaths) {
        if (Test-Path $configPath) {
            try {
                # Check InstalledLanguages property
                $installedLanguages = Get-ItemProperty -Path $configPath -Name "InstalledLanguages" -ErrorAction SilentlyContinue
                if ($installedLanguages -and $installedLanguages.InstalledLanguages) {
                    Write-DetectionLog "Found InstalledLanguages: $($installedLanguages.InstalledLanguages)"
                    if ($installedLanguages.InstalledLanguages -match $Lang) {
                        Write-DetectionLog "Proofing Tools for $Lang found in InstalledLanguages"
                        $Installed = $true
                        break
                    }
                }
                
                # Check ProductReleaseIds for proofing tools
                $productIds = Get-ItemProperty -Path $configPath -Name "ProductReleaseIds" -ErrorAction SilentlyContinue
                if ($productIds -and $productIds.ProductReleaseIds -match "ProofingTools") {
                    Write-DetectionLog "ProofingTools product detected in ProductReleaseIds"
                    # If ProofingTools product exists and language is in InstalledLanguages, it's installed
                    if ($installedLanguages -and $installedLanguages.InstalledLanguages -match $Lang) {
                        $Installed = $true
                        break
                    }
                }
            }
            catch {
                Write-DetectionLog "Warning: Could not read Click-to-Run configuration: $($_.Exception.Message)"
            }
        }
    }

    # Method 2: Check Uninstall registry keys
    if (-not $Installed) {
        Write-DetectionLog "Method 2: Checking Uninstall registry keys..."
        $UninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )

        foreach ($path in $UninstallPaths) {
            if (Test-Path $path) {
                try {
                    # Check for specific proofing tools registry key patterns
                    $proofingKeyPatterns = @(
                        "O365ProPlusRetail - $Lang.proof",
                        "O365BusinessRetail - $Lang.proof",
                        "ProofingTools_$Lang",
                        "*ProofingTools*$Lang*"
                    )

                    foreach ($proofPattern in $proofingKeyPatterns) {
                        # Handle wildcard patterns
                        if ($proofPattern -like "*`**") {
                            $keys = Get-ChildItem $path -ErrorAction SilentlyContinue
                            foreach ($k in $keys) {
                                $keyName = $k.PSChildName
                                if ($keyName -like $proofPattern) {
                                    Write-DetectionLog "Found proofing tools registry key (wildcard): $keyName"
                                    $Installed = $true
                                    break
                                }
                            }
                        }
                        else {
                            $proofingKeyPath = Join-Path $path $proofPattern
                            if (Test-Path $proofingKeyPath) {
                                Write-DetectionLog "Found proofing tools registry key: $proofPattern"
                                $Installed = $true
                                break
                            }
                        }
                    }
                    
                    if ($Installed) { break }

                    # Fallback: Check all keys for proofing tools pattern
                    if (-not $Installed) {
                        $keys = Get-ChildItem $path -ErrorAction SilentlyContinue
                        foreach ($k in $keys) {
                            $keyName = $k.PSChildName
                            # Match patterns like: O365ProPlusRetail - nl-nl.proof or O365BusinessRetail - nl-nl.proof
                            if ($keyName -match "O365(ProPlusRetail|BusinessRetail)\s*-\s*$Lang\.proof") {
                                Write-DetectionLog "Found proofing tools registry key (regex): $keyName"
                                $Installed = $true
                                break
                            }
                        }
                    }
                }
                catch {
                    Write-DetectionLog "Warning: Could not read uninstall registry path $path : $($_.Exception.Message)"
                }
            }
            if ($Installed) { break }
        }
    }

    # Method 3: Check for proofing tools files in Office installation directory
    if (-not $Installed) {
        Write-DetectionLog "Method 3: Checking Office installation directory for proofing files..."
        $OfficePaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\Proof\$Lang",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Proof\$Lang"
        )

        foreach ($officePath in $OfficePaths) {
            if (Test-Path $officePath) {
                # Check if the directory contains proofing files
                $proofFiles = Get-ChildItem -Path $officePath -Filter "*.dll" -ErrorAction SilentlyContinue
                if ($proofFiles -and $proofFiles.Count -gt 0) {
                    Write-DetectionLog "Found proofing tools files in: $officePath"
                    $Installed = $true
                    break
                }
            }
        }
    }

    return $Installed
}

# Main execution
try {
    Write-DetectionLog "=== Microsoft 365 Proofing Tools Detection Started ==="
    Write-DetectionLog "Target Language: $LanguageID"
    
    if (Test-ProofingToolsInstalled -Lang $LanguageID) {
        Write-DetectionLog "=== DETECTION SUCCESS: Proofing Tools for $LanguageID is installed ==="
        Write-Output "Proofing Tools for $LanguageID is installed."
        exit 0  # Detection success
    } else {
        Write-DetectionLog "=== DETECTION FAILED: Proofing Tools for $LanguageID is NOT installed ==="
        Write-Output "Proofing Tools for $LanguageID is NOT installed."
        exit 1  # Detection failed
    }
}
catch {
    Write-DetectionLog "=== DETECTION ERROR: $($_.Exception.Message) ==="
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}