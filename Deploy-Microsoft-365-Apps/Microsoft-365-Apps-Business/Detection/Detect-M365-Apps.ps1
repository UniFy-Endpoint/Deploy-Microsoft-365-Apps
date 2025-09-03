# Microsoft 365 Apps detection (no version check)
# Returns 0 if detected, 1 if not detected
# Supports both AMD64 (x64) and ARM64

# --- Editable criteria ---
# Product ID to detect (matches your Configuration.xml)
$RequireProductId     = 'O365BusinessRetail'   # e.g., O365ProPlusRetail, O365BusinessRetail, etc.
# Acceptable architectures reported by ClickToRun 'Platform'. Set to $null to ignore architecture.
$AllowedArchitectures = @('x64','arm64')      # Accept both AMD64 and ARM64
# -------------------------

$ErrorActionPreference = 'Stop'

function Write-Dbg($msg) { Write-Output $msg }

function Get-OfficeC2RConfig {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
  )
  foreach ($p in $paths) {
    try {
      $prop = Get-ItemProperty -Path $p -ErrorAction Stop
      if ($prop) {
        [pscustomobject]@{
          Path              = $p
          ProductReleaseIds = $prop.ProductReleaseIds
          Platform          = $prop.Platform
        }
      }
    } catch {}
  }
}

try {
  $configs = @(Get-OfficeC2RConfig)
  if (-not $configs -or $configs.Count -eq 0) {
    Write-Dbg "Office ClickToRun configuration registry not found"
    exit 1
  }

  foreach ($cfg in $configs) {
    # Product IDs are comma-separated in ProductReleaseIds
    $products = @()
    if ($cfg.ProductReleaseIds) {
      $products = ($cfg.ProductReleaseIds -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    if ($RequireProductId -and ($products -notcontains $RequireProductId)) {
      Write-Dbg "Path: $($cfg.Path) - Required product '$RequireProductId' not found in: $($products -join ', ')"
      continue
    }

    # Architecture match (optional)
    $arch = $cfg.Platform
    if ($AllowedArchitectures -and $arch) {
      $allowed = $AllowedArchitectures | ForEach-Object { $_.ToLower() }
      if (-not ($allowed -contains $arch.ToLower())) {
        Write-Dbg "Path: $($cfg.Path) - Architecture '$arch' not in allowed set: $($AllowedArchitectures -join ', ')"
        continue
      }
    }

    # If we reach here, detection succeeded
    Write-Dbg "Detected Microsoft 365 Apps ($RequireProductId). Path=$($cfg.Path) Arch=$arch"
    exit 0
  }

  # None matched
  Write-Dbg "Microsoft 365 Apps ($RequireProductId) not detected with required criteria."
  exit 1
}
catch {
  # Return non-zero so Intune can retry later
  Write-Dbg "Detection error: $($_.Exception.Message)"
  exit 1
}