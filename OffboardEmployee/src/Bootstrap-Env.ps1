<# ==============================================================================================
Bootstrap-Env.ps1
--------------------------------------------------------------------------------------------------
Goal
----
Standardize our local PowerShell environment: enforce PowerShell 7 (x64) and install
the right admin modules.

What it does
------------
• Confirms we are in PowerShell 7 (x64) and at least version 7.4.0.
• Trusts the PowerShell Gallery (for installs) if needed.
• Installs/loads: ExchangeOnlineManagement and Microsoft.Graph (minimum versions).
================================================================================================ #>

[CmdletBinding()]
param([switch]$Force)

function Say([string]$m){ Write-Host $m }
function Fail([string]$m){ throw $m }

# Check PowerShell edition and bitness
if (-not [Environment]::Is64BitProcess) {
  Fail "You are running a 32‑bit PowerShell host. Please start 'PowerShell 7 (x64)'."
}
if ($PSVersionTable.PSVersion.Major -lt 7 -or $PSVersionTable.PSVersion -lt [version]'7.4.0') {
  Fail "PowerShell 7.4+ required. Winget: winget install --id Microsoft.PowerShell"
}

# Trust PSGallery if needed
try {
  $pol = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
  if ($pol.InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
  }
} catch {
  Register-PSRepository -Default
  Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
}

# Install modules
$mods = @(
  @{ Name='ExchangeOnlineManagement'; Min='3.3.0' },
  @{ Name='Microsoft.Graph';         Min='2.16.0' }
)
foreach ($m in $mods) {
  $have = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
  if (-not $have -or ($have.Version -lt [Version]$m.Min) -or $Force) {
    try { Install-Module $m.Name -Force -Scope AllUsers -MinimumVersion $m.Min -ErrorAction Stop }
    catch { Install-Module $m.Name -Force -Scope CurrentUser -MinimumVersion $m.Min -ErrorAction Stop }
  }
  Import-Module $m.Name -ErrorAction Stop
}
Say "Environment looks good. PowerShell $($PSVersionTable.PSVersion) x64; modules loaded."
