<# ==============================================================================================
Bootstrap-Env.ps1
--------------------------------------------------------------------------------------------------
Goal
- Set up a working PowerShell environment for the offboarding script.
- Install required PowerShell modules.

This script is safe to run more than once.
It only installs what is missing or out of date.

Notes
- This script is meant for Windows.
- It prefers installing modules for the current user so it works without administrator rights.
- If you run PowerShell as administrator, you can choose to install for all users.
============================================================================================== #>

[CmdletBinding()]
param(
  [switch]$Force,
  [switch]$InstallForAllUsers,
  [switch]$InstallActiveDirectoryTools
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Say([string]$m){ Write-Host $m }
function Warn([string]$m){ Write-Warning $m }
function Fail([string]$m){ throw $m }

# Allow this PowerShell window only to run the setup steps.
# This resets when you close the window.
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}

# Check we are in 64 bit PowerShell.
if (-not [Environment]::Is64BitProcess) {
  Fail "You are running 32 bit PowerShell. Please open PowerShell 7 (64 bit) and try again."
}

# Check PowerShell version.
# Offboarding tooling expects PowerShell 7 or newer.
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Warn "PowerShell 7 is required. If you do not have it, install it from the Microsoft Store, the Company Portal, or run: winget install --id Microsoft.PowerShell"
  Fail "PowerShell 7 is not installed in this window."
}

# Optional: require a minimum version if you want stricter control.
$minPs = [version]'7.4.0'
if ($PSVersionTable.PSVersion -lt $minPs) {
  Warn ("Your PowerShell version is {0}. We recommend {1} or newer." -f $PSVersionTable.PSVersion, $minPs)
}

# Use modern security protocol settings when possible.
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

# Ensure NuGet provider exists.
try {
  $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
  if (-not $nuget) {
    Say "Installing NuGet package provider..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
  }
} catch {
  Warn ("NuGet provider setup had an issue: {0}" -f $_)
}

# Trust PowerShell Gallery so installs can happen without prompts.
try {
  $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
  if ($repo.InstallationPolicy -ne 'Trusted') {
    Say "Setting PowerShell Gallery to Trusted..."
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
  }
} catch {
  Say "Registering PowerShell Gallery..."
  Register-PSRepository -Default
  Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
}

function Install-OrUpdateModule {
  param(
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][version]$MinimumVersion
  )

  $have = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1

  $needInstall = $Force -or (-not $have) -or ($have.Version -lt $MinimumVersion)

  if ($needInstall) {
    Say ("Installing or updating module {0} (minimum {1})" -f $Name, $MinimumVersion)

    $scope = if ($InstallForAllUsers -and ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
      'AllUsers'
    } else {
      'CurrentUser'
    }

    try {
      Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope $scope -Force -AllowClobber -ErrorAction Stop
    } catch {
      # Fallback: try CurrentUser if AllUsers failed.
      if ($scope -ne 'CurrentUser') {
        Warn "All users install failed. Trying current user install."
        Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
      } else {
        throw
      }
    }
  } else {
    Say ("Module {0} is already installed (version {1})" -f $Name, $have.Version)
  }

  Import-Module $Name -ErrorAction Stop
}

# Required modules for the offboarding script.
# These minimum versions match what the offboarding script expects.
Install-OrUpdateModule -Name 'ExchangeOnlineManagement' -MinimumVersion ([version]'3.7.2')
Install-OrUpdateModule -Name 'Microsoft.Graph' -MinimumVersion ([version]'2.16.0')

# Optional: install Active Directory tools.
# These are only needed if you plan to use the on premises options in the offboarding script.
if ($InstallActiveDirectoryTools) {
  if (-not $IsWindows) {
    Warn "Active Directory tools can only be installed on Windows. Skipping."
  } else {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
      Warn "Active Directory tools installation requires administrator rights. Run PowerShell as administrator and try again."
    } else {
      try {
        Say "Installing Active Directory tools..."
        # This feature name works on Windows 10 and Windows 11.
        Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' | Out-Null
        Say "Active Directory tools installed."
      } catch {
        Warn ("Active Directory tools install failed: {0}" -f $_)
      }
    }
  }
}

Say ("Environment looks good. PowerShell version: {0}. Modules loaded." -f $PSVersionTable.PSVersion)
