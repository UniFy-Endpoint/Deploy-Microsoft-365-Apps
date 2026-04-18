<#
.SYNOPSIS
    Detection Script for Microsoft 365 Language Packs

.DESCRIPTION
    This script is meant to be used with Intune (Win32 app deployment).
    It checks whether the specified Microsoft 365 Language Pack is installed.
    Detection is based on Office Click-to-Run configuration and presence of language pack registry keys.

.PARAMETER LanguageID
    Define the Language Pack ID/Tag (e.g. "en-us", "nl-nl", "de-de")
    
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

function Test-LanguagePackInstalled {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Lang
    )

    $Installed = $false
    Write-DetectionLog "Starting Language Pack detection for language: $Lang"

    # Method 1: Check Office Click-to-Run Configuration registry
    Write-DetectionLog "Method 1: Checking Office Click-to-Run Configuration..."
    $OfficeConfigPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )

    foreach ($configPath in $OfficeConfigPaths) {
        if (Test-Path $configPath) {
            try {
                # Check InstalledLanguages property (primary method)
                $installedLanguages = Get-ItemProperty -Path $configPath -Name "InstalledLanguages" -ErrorAction SilentlyContinue
                if ($installedLanguages -and $installedLanguages.InstalledLanguages) {
                    Write-DetectionLog "Found InstalledLanguages: $($installedLanguages.InstalledLanguages)"
                    if ($installedLanguages.InstalledLanguages -match $Lang) {
                        Write-DetectionLog "Language Pack for $Lang found in InstalledLanguages"
                        $Installed = $true
                        break
                    }
                }
                
                # Check product-specific language properties
                $productLanguageProperties = @(
                    "O365ProPlusRetail.InstallLanguage",
                    "O365BusinessRetail.InstallLanguage",
                    "ProPlusRetail.InstallLanguage"
                )
                
                foreach ($prop in $productLanguageProperties) {
                    $langProp = Get-ItemProperty -Path $configPath -Name $prop -ErrorAction SilentlyContinue
                    if ($langProp -and $langProp.$prop -match $Lang) {
                        Write-DetectionLog "Language Pack $Lang found in property: $prop"
                        $Installed = $true
                        break
                    }
                }
                
                if ($Installed) { break }
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
                    # Check for specific Language Pack registry key patterns
                    $LanguagePackKeyPatterns = @(
                        "O365ProPlusRetail - $Lang",
                        "O365BusinessRetail - $Lang",
                        "LanguagePack_$Lang",
                        "Microsoft 365 Apps for enterprise - $Lang",
                        "Microsoft Office Professional Plus 2019 - $Lang"
                    )

                    foreach ($LanguagePack in $LanguagePackKeyPatterns) {
                        $LanguagePackKeyPath = Join-Path $path $LanguagePack
                        if (Test-Path $LanguagePackKeyPath) {
                            Write-DetectionLog "Language Pack registry key found: $LanguagePack"
                            $Installed = $true
                            break
                        }
                    }
                    
                    if ($Installed) { break }

                    # Fallback: Check all keys for Language Pack pattern
                    $keys = Get-ChildItem $path -ErrorAction SilentlyContinue
                    foreach ($k in $keys) {
                        $keyName = $k.PSChildName
                        # Match patterns for language packs (not proofing tools)
                        if (($keyName -match "O365(ProPlusRetail|BusinessRetail)\s*-\s*$Lang$" -or 
                             $keyName -match "Microsoft\s+(365|Office).*-\s*$Lang$" -or
                             $keyName -match "LanguagePack.*$Lang") -and 
                            $keyName -notmatch "\.proof") {
                            Write-DetectionLog "Found Language Pack registry key (regex): $keyName"
                            $Installed = $true
                            break
                        }
                    }
                    
                    if ($Installed) { break }
                }
                catch {
                    Write-DetectionLog "Warning: Could not read uninstall registry path $path : $($_.Exception.Message)"
                }
            }
        }
    }

    # Method 3: Check for language pack files in Office installation directory
    if (-not $Installed) {
        Write-DetectionLog "Method 3: Checking Office installation directory for language files..."
        $OfficePaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\$Lang",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\$Lang"
        )

        foreach ($officePath in $OfficePaths) {
            if (Test-Path $officePath) {
                # Check if the directory contains language resource files
                $langFiles = Get-ChildItem -Path $officePath -Filter "*.dll" -ErrorAction SilentlyContinue
                if ($langFiles -and $langFiles.Count -gt 0) {
                    Write-DetectionLog "Found language pack files in: $officePath"
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
    Write-DetectionLog "=== Microsoft 365 Language Pack Detection Started ==="
    Write-DetectionLog "Target Language: $LanguageID"
    
    if (Test-LanguagePackInstalled -Lang $LanguageID) {
        Write-DetectionLog "=== DETECTION SUCCESS: Language Pack for $LanguageID is installed ==="
        Write-Output "SUCCESS: Language Pack for $LanguageID is installed."
        exit 0  # Detection success
    } else {
        Write-DetectionLog "=== DETECTION FAILED: Language Pack for $LanguageID is NOT installed ==="
        Write-Output "NOT FOUND: Language Pack for $LanguageID is NOT installed."
        exit 1  # Detection failed
    }
}
catch {
    Write-DetectionLog "=== DETECTION ERROR: $($_.Exception.Message) ==="
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}