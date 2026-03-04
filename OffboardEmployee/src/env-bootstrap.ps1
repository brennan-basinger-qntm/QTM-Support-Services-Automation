<# ==============================================================================================
Bootstrap-Env.ps1

Goal
- Set up a working PowerShell environment for the offboarding script.
- Install required PowerShell modules.
- Optional: install Windows Active Directory tools.

Notes
- This script is meant for Windows.
- It prefers installing modules for the current user so it works without administrator rights.
- It only changes execution policy for the current PowerShell process (this window only).
============================================================================================== #>

[CmdletBinding()]
param(
  [switch]$Force,
  [switch]$InstallForAllUsers,
  [switch]$InstallActiveDirectoryTools,
  [switch]$TrustPowerShellGallery,
  [version]$MinimumPowerShellVersion = [version]'7.4.0',
  [version]$MinimumPowerShellGetVersion = [version]'2.2.5',
  [version]$MinimumPackageManagementVersion = [version]'1.4.8.1'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Say([string]$m){ Write-Host $m }
function Warn([string]$m){ Write-Warning $m }
function Fail([string]$m){ throw $m }

# Allow this PowerShell window only to run the setup steps.
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}

# Check we are in 64 bit PowerShell.
if (-not [Environment]::Is64BitProcess) {
  Fail "You are running 32 bit PowerShell. Please open PowerShell 7 (64 bit) and try again."
}

# Check PowerShell version.
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Warn "PowerShell 7 is required. Install it from the Microsoft Store or the Company Portal."
  Warn "If your organization allows it, you can also run: winget install --id Microsoft.PowerShell"
  Fail "PowerShell 7 is not installed in this window."
}

if ($PSVersionTable.PSVersion -lt $MinimumPowerShellVersion) {
  Warn ("Your PowerShell version is {0}. We recommend {1} or newer." -f $PSVersionTable.PSVersion, $MinimumPowerShellVersion)
}

# Use modern encryption settings when possible for downloads.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

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

# Ensure PowerShellGet and PackageManagement are healthy (reduces install failures).
function Ensure-ModuleFramework {
  param(
    [version]$MinPowerShellGet,
    [version]$MinPackageManagement
  )

  $psGet = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
  if (-not $psGet -or $psGet.Version -lt $MinPowerShellGet) {
    Say ("Installing or updating PowerShellGet (minimum {0})..." -f $MinPowerShellGet)
    Install-Module -Name PowerShellGet -MinimumVersion $MinPowerShellGet -Scope CurrentUser -Force -AllowClobber
  }

  $pkgMgmt = Get-Module -ListAvailable -Name PackageManagement | Sort-Object Version -Descending | Select-Object -First 1
  if (-not $pkgMgmt -or $pkgMgmt.Version -lt $MinPackageManagement) {
    Say ("Installing or updating PackageManagement (minimum {0})..." -f $MinPackageManagement)
    Install-Module -Name PackageManagement -MinimumVersion $MinPackageManagement -Scope CurrentUser -Force -AllowClobber
  }
}

Ensure-ModuleFramework -MinPowerShellGet $MinimumPowerShellGetVersion -MinPackageManagement $MinimumPackageManagementVersion

# Register and optionally trust PowerShell Gallery.
try {
  $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
} catch {
  Say "Registering PowerShell Gallery..."
  Register-PSRepository -Default
  $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
}

if ($TrustPowerShellGallery) {
  if ($repo.InstallationPolicy -ne 'Trusted') {
    Say "Setting PowerShell Gallery to Trusted..."
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
  }
} else {
  Say "PowerShell Gallery trust was not changed. If you get prompts during installs, rerun with -TrustPowerShellGallery."
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

    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).
      IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    $scope =
      if ($InstallForAllUsers -and $isAdmin) { 'AllUsers' }
      else { 'CurrentUser' }

    try {
      Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope $scope -Force -AllowClobber -ErrorAction Stop
    } catch {
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
Install-OrUpdateModule -Name 'ExchangeOnlineManagement' -MinimumVersion ([version]'3.7.2')
Install-OrUpdateModule -Name 'Microsoft.Graph' -MinimumVersion ([version]'2.16.0')

# Optional: install Windows Active Directory tools.
if ($InstallActiveDirectoryTools) {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).
    IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

  if (-not $isAdmin) {
    Warn "Active Directory tools installation requires administrator rights. Run PowerShell as administrator and try again."
  } else {
    try {
      Say "Installing Active Directory tools..."
      Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' | Out-Null
      Say "Active Directory tools installed."
    } catch {
      Warn ("Active Directory tools install failed: {0}" -f $_)
    }
  }
}

# Summary
$exo = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
$graph = Get-Module -ListAvailable -Name Microsoft.Graph | Sort-Object Version -Descending | Select-Object -First 1
$ad = Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1

Say ""
Say ("Environment looks good.")
Say ("PowerShell version: {0}" -f $PSVersionTable.PSVersion)
Say ("ExchangeOnlineManagement: {0}" -f ($(if($exo){$exo.Version}else{'Not found'})))
Say ("Microsoft.Graph: {0}" -f ($(if($graph){$graph.Version}else{'Not found'})))
Say ("ActiveDirectory module: {0}" -f ($(if($ad){'Available'}else{'Not installed'})))