<# =================================================================================================
Invoke-UserOffboarding.ps1
----------------------------------------------------------------------------------------------------
Offboarding runbook automation for Microsoft 365 / Entra ID using:
  • Exchange Online PowerShell (EXO)
  • Microsoft Graph PowerShell (Mg)

SAFE BY DEFAULT — the script runs in Preview unless you pass -Apply.
It captures BEFORE / AFTER snapshots and writes paste-ready ServiceNow work notes.

USAGE (Preview)
  .\Invoke-UserOffboarding.ps1 -UserUpn first.last@quantinuum.com -TicketNumber INC12345678

USAGE (Apply)
  .\Invoke-UserOffboarding.ps1 -UserUpn first.last@quantinuum.com -TicketNumber INC12345678 -Apply

NOTES
  • Distribution lists (DGs) & mail-enabled security groups are removed via EXO.
  • Microsoft 365/Security groups are removed via Graph.
  • Dynamic groups are detected and **never** changed (we only list them).
  • Mailbox is converted to Shared by default and stamped with a future expiry marker
    in CustomAttribute15 - place to store metadata and formatted like (e.g., "Expires: 2026-04-21 (180d)").
  • AD/on‑prem steps are optional and skipped unless you request them AND the AD module is available.

================================================================================================ #>

[CmdletBinding()]
param(
  # Core
  [Parameter(Mandatory=$true)][string]$UserUpn,
  [Parameter(Mandatory=$true)][string]$TicketNumber,

  # Mailbox handling
  [switch]$ConvertMailboxToShared = $true,
  [int]$SharedMailboxExpiryDays = 180,

  # Supervisor / manager options
  [string]$SupervisorUpn,
  [string]$BackupOwnerUpn,
  [switch]$GrantSupervisorFullAccess,
  [switch]$GrantSupervisorSendAs,

  # Group & license cleanup
  [switch]$RemoveFromDistributionLists = $true,
  [switch]$RemoveFromGroups          = $true,
  [switch]$RemoveMailboxDelegations  = $true,
  [switch]$RemoveLicenses = $true,
  [switch]$DisableEntraSignIn = $true,

  # Active Directory (on‑prem) — optional. If not available, we skip.
  [switch]$DisableAD,
  [switch]$UpdateAdDescription,
  [string]$DisabledOuDn,                      # e.g. "OU=Disabled Users,OU=Corp,DC=contoso,DC=com"

  # Execution control
  [switch]$Apply,                             # do changes only when present
  [string]$TenantHint = '94c4857e-1130-4ab8-8eac-069b40c9db20',                        # optional tenant id or verified domain for Graph
  [switch]$UseElevatedGraphScopes,            # adds Directory.ReadWrite.All
  [string]$OutputFolder = (Join-Path $env:USERPROFILE ("Desktop\Offboarding-" + (Get-Date -Format "yyyyMMdd-HHmmss")))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------- helpers ----------
function Say([string]$msg){ Write-Host $msg }
function Step([string]$msg){ Write-Host "== $msg" -ForegroundColor Cyan }
function Act ([string]$msg){ Write-Host $msg -ForegroundColor Yellow }
function Skip([string]$msg){ Write-Warning $msg }
function Did ([string]$msg){ Write-Host $msg -ForegroundColor Green }

$Preview = -not $Apply

# Create output folder & transcript
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$TranscriptPath = Join-Path $OutputFolder ("Transcript-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt")
Start-Transcript -Path $TranscriptPath -Append | Out-Null

# ---------- environment checks ----------
function Ensure-ModuleLoaded {
  param([string]$Name,[Version]$MinVersion)
  $have = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
  if (-not $have -or ($MinVersion -and $have.Version -lt $MinVersion)) {
    Act "Installing module $Name (min $MinVersion)"
    try { Install-Module $Name -Force -Scope AllUsers -MinimumVersion $MinVersion -ErrorAction Stop }
    catch { Install-Module $Name -Force -Scope CurrentUser -MinimumVersion $MinVersion -ErrorAction Stop }
  }
  Import-Module $Name -ErrorAction Stop
}

# Now passing organization via Tenant ID to Graph
function Ensure-EXO {
  Ensure-ModuleLoaded -Name ExchangeOnlineManagement -MinVersion ([Version]'3.3.0')
  if (-not (Get-ConnectionInformation)) {
    Act "Connecting to Exchange Online..."
    $isDomain = ($TenantHint -and ($TenantHint -match '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'))
    if ($isDomain) {
      Connect-ExchangeOnline -ShowBanner:$false -Organization $TenantHint -ErrorAction Stop | Out-Null
    } else {
      Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
    }
  }
}


function Ensure-Graph {
  param([string[]]$Scopes)
  Ensure-ModuleLoaded -Name Microsoft.Graph -MinVersion ([Version]'2.16.0')
  try {
    $ctx = (Get-MgContext) 2>$null
    $need = $true
    if ($ctx) {
      # Make sure we have all required scopes
      $haveScopes = @($ctx.Scopes)
      $missing = @($Scopes | Where-Object { $_ -notin $haveScopes })
      if ($missing.Count -eq 0) { $need = $false }
    }
    if ($need) {
      Act "Connecting to Microsoft Graph with scopes: $($Scopes -join ', ')"
      if ($TenantHint) {
        Connect-MgGraph -Scopes $Scopes -TenantId $TenantHint -NoWelcome | Out-Null
      } else {
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
      }
    }
  } catch {
    throw "Failed to connect to Microsoft Graph: $_"
  }
}

function Ensure-ADLocal {
  $ad = Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1
  if ($ad) { Import-Module ActiveDirectory -ErrorAction Stop; return $true }
  return $false
}

# ---------- utility ----------
function Resolve-GraphUser {
  param([Parameter(Mandatory=$true)][string]$Identity)
  try {
    # Supports UPN, objectId, email
    $u = Get-MgUser -UserId $Identity -Property "id,userPrincipalName,displayName,mail" -ErrorAction Stop
  } catch {
    $u = Get-MgUser -Filter "userPrincipalName eq '$Identity'" -Property "id,userPrincipalName,displayName,mail" -ErrorAction SilentlyContinue
    if (-not $u) { $u = Get-MgUser -Filter "mail eq '$Identity'" -Property "id,userPrincipalName,displayName,mail" -ErrorAction SilentlyContinue }
  }
  if (-not $u) { throw "Cannot resolve user '$Identity' in Graph." }
  return $u
}

# ---------- connect services ----------
$graphScopes = @('User.ReadWrite.All','Group.ReadWrite.All','GroupMember.ReadWrite.All','Directory.Read.All','AuditLog.Read.All')
if ($UseElevatedGraphScopes) { $graphScopes += 'Directory.ReadWrite.All' }
Ensure-Graph -Scopes $graphScopes
Ensure-EXO

# ---------- locate the user ----------
Step "Locating user '$UserUpn'"
try {
  $User = Resolve-GraphUser -Identity $UserUpn
} catch { Stop-Transcript | Out-Null; throw }



# ---------- snapshot helpers ----------
function Snapshot-GraphGroups {
  param([string]$UserId)
  $out = @()

  # Direct group memberships only
  # If we ever want transitive groups included, we can switch to Get-MgUserTransitiveMemberOfAsGroup -All.
  $groups = Get-MgUserMemberOfAsGroup -UserId $UserId -All -ErrorAction SilentlyContinue

  foreach ($g in $groups) {
    try {
      # Pull properties needed to detect dynamic groups & categorize the type
      $gg = Get-MgGroup -GroupId $g.Id -Property "id,displayName,groupTypes,securityEnabled,mail,mailEnabled,membershipRule,membershipRuleProcessingState" -ErrorAction SilentlyContinue
      if ($gg) {
        $isDynamic = -not [string]::IsNullOrEmpty($gg.membershipRule) -or ($gg.groupTypes -contains 'DynamicMembership')
        $isUnified = ($gg.groupTypes -contains 'Unified')
        $out += [pscustomobject]@{
          GroupId     = $gg.Id
          DisplayName = $gg.DisplayName
          Mail        = $gg.Mail
          MailEnabled = [bool]$gg.MailEnabled
          IsSecurity  = [bool]$gg.SecurityEnabled
          IsUnified   = $isUnified
          IsDynamic   = $isDynamic
        }
      }
    } catch {
      Skip "Failed to read group $($g.Id): $_"
    }
  }
  return $out | Sort-Object DisplayName
}



function Snapshot-GraphOwnedGroups {
  param([string]$UserId)
  $out = @()
  try {
    $ownedGroups = Get-MgUserOwnedObjectAsGroup -UserId $UserId -All -ErrorAction SilentlyContinue
    foreach ($g in $ownedGroups) {
      try {
        # Basic properties for display/categorization
        $gg = Get-MgGroup -GroupId $g.Id -Property "id,displayName,groupTypes" -ErrorAction SilentlyContinue
        if ($gg) {
          # Count owners (works for M365/security groups but not for classic Exchange DLs)
          $owners = Get-MgGroupOwner -GroupId $gg.Id -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }
          $out += [pscustomobject]@{
            GroupId     = $gg.Id
            DisplayName = $gg.DisplayName
            OwnersCount = @($owners).Count
            IsUnified   = ($gg.GroupTypes -contains 'Unified')
          }
        }
      } catch {
        Skip "Failed to resolve owned group $($g.Id): $_"
      }
    }
  } catch { }
  return $out
}



function Snapshot-EXO-DLs {
  param([string]$UserSmtp)
  $dlMatches = @()
  $dls = Get-DistributionGroup -ResultSize Unlimited
  foreach ($dl in $dls) {
    $isDynamic = $dl.RecipientTypeDetails -eq 'DynamicDistributionGroup'
    try {
      # (Dynamic DLs don't support Get-DistributionGroupMember; treat as dynamic and only list)
      if (-not $isDynamic) {
        $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue
        if ($members | Where-Object { $_.PrimarySmtpAddress -ieq $UserSmtp }) {
          $dlMatches += [pscustomobject]@{
            DisplayName = $dl.DisplayName
            PrimarySmtp = $dl.PrimarySmtpAddress
            IsDynamic   = $false
          }
        }
      } else {
        # For dynamic DGs, check approximate membership (Exchange doesn't easily filter by user)
        # We only list the DG, not confirm membership, to avoid heavy evaluation.
        $dlMatches += [pscustomobject]@{
          DisplayName = $dl.DisplayName
          PrimarySmtp = $dl.PrimarySmtpAddress
          IsDynamic   = $true
        }
      }
    } catch {
      Skip "DL scan failed for $($dl.DisplayName): $_"
    }
  }
  return $dlMatches | Sort-Object DisplayName
}

function Snapshot-EXO-Delegations {
  param([string]$UserSmtp)
  $out = @()
  try {
    $mbx = Get-Mailbox -Identity $UserSmtp -ErrorAction Stop
  } catch {
    return @()
  }

  # FullAccess
  try {
    $fa = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
          Where-Object { -not $_.IsInherited -and $_.User -notmatch 'NT AUTHORITY\\SELF' -and $_.AccessRights -contains 'FullAccess' }
    foreach ($p in $fa) {
      $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='FullAccess'; Trustee=$p.User }
    }
  } catch { }

  # SendAs
  try {
    $sa = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
          Where-Object { -not $_.IsInherited -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'SendAs' }
    foreach ($p in $sa) {
      $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='SendAs'; Trustee=$p.Trustee }
    }
  } catch { }

  # SendOnBehalf
  try {
    $sob = (Get-Mailbox -Identity $mbx.Identity -Property GrantSendOnBehalfTo -ErrorAction SilentlyContinue).GrantSendOnBehalfTo
    foreach ($t in $sob) {
      $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='SendOnBehalf'; Trustee=$t.PrimarySmtpAddress }
    }
  } catch { }

  return $out | Sort-Object Mailbox, Right, Trustee
}

function Snapshot-Licenses {
  param([string]$UserId)
  try {
    $lic = Get-MgUserLicenseDetail -UserId $UserId -All
    return $lic | ForEach-Object {
      [pscustomobject]@{
        SkuId   = $_.SkuId
        SkuPart = $_.SkuPartNumber
        Svc     = ($_.ServicePlans | Where-Object { $_.ProvisioningStatus -eq 'Success' } | Select-Object -ExpandProperty ServicePlanName) -join ';'
      }
    }
  } catch {
    Skip "License snapshot failed: $_"
    return @()
  }
}

function Snapshot-ADGroups {
  param([string]$SamAccountName)
  $out = @()
  try {
    $adUser = Get-ADUser -Identity $SamAccountName -Properties MemberOf,Description,Enabled -ErrorAction Stop
    $out += [pscustomobject]@{
      AD_Enabled    = $adUser.Enabled
      AD_Description= $adUser.Description
    }
    foreach ($dn in $adUser.MemberOf) {
      try {
        $g = Get-ADGroup -Identity $dn -ErrorAction SilentlyContinue
        if ($g) { $out += [pscustomobject]@{ GroupName=$g.Name; DistinguishedName=$g.DistinguishedName } }
      } catch { }
    }
  } catch {
    Skip "AD snapshot skipped (module not available or user not found)."
  }
  return $out
}

# ---------- BEFORE snapshot ----------
Step "Snapshot BEFORE"
$Before = [ordered]@{
  Identity = @{ DisplayName=$User.DisplayName; UPN=$User.UserPrincipalName; Id=$User.Id; Mail=$User.Mail }
  EXO      = [ordered]@{}
  Graph    = [ordered]@{}
  Licenses = @()
  AD       = @()
}

# Check mailbox
try {
  $mbx = Get-Mailbox -Identity $UserUpn -ErrorAction Stop
  $Before.EXO.Mailbox = @{
    PrimarySmtp             = $mbx.PrimarySmtpAddress
    RecipientTypeDetails    = $mbx.RecipientTypeDetails
    CustomAttribute15       = $mbx.CustomAttribute15
  }
} catch { $mbx = $null; $Before.EXO.Mailbox = $null }

# Gather memberships and delegations
$Before.EXO.DLs         = Snapshot-EXO-DLs        -UserSmtp $UserUpn
$Before.EXO.Delegations = Snapshot-EXO-Delegations -UserSmtp $UserUpn
$Before.Graph.Groups    = Snapshot-GraphGroups     -UserId   $User.Id
$Before.Graph.Owns     = Snapshot-GraphOwnedGroups -UserId   $User.Id
$Before.Licenses        = Snapshot-Licenses        -UserId   $User.Id

# Optional AD snapshot
$HaveAD = $false
if ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn) {
  $HaveAD = Ensure-ADLocal
  if ($HaveAD) {
    # Try SAM from UPN
    $sam = ($User.UserPrincipalName -split '@')[0]
    $Before.AD = Snapshot-ADGroups -SamAccountName $sam
  } else {
    Skip "AD module not available locally — AD steps will be skipped."
  }
}

# Write BEFORE snapshots
$Before.EXO.DLs         | Export-Csv -Path (Join-Path $OutputFolder 'Before-EXO-DLs.csv') -NoTypeInformation -Encoding UTF8
$Before.EXO.Delegations | Export-Csv -Path (Join-Path $OutputFolder 'Before-EXO-Delegations.csv') -NoTypeInformation -Encoding UTF8
$Before.Graph.Groups    | Export-Csv -Path (Join-Path $OutputFolder 'Before-Graph-Groups.csv') -NoTypeInformation -Encoding UTF8
$Before.Licenses        | Export-Csv -Path (Join-Path $OutputFolder 'Before-Licenses.csv') -NoTypeInformation -Encoding UTF8
if ($Before.AD) { $Before.AD | Export-Csv -Path (Join-Path $OutputFolder 'Before-AD.csv') -NoTypeInformation -Encoding UTF8 }

# ---------- build plan ----------
$Plan = New-Object System.Collections.Generic.List[object]

function Add-Plan { param([string]$Area,[string]$Action) $Plan.Add([pscustomobject]@{ Area=$Area; Action=$Action }) }

$willConvert = $ConvertMailboxToShared -and $mbx -and $mbx.RecipientTypeDetails -notlike '*SharedMailbox*'
if ($willConvert) { Add-Plan 'Mailbox'  "Convert mailbox to Shared and stamp expiry +$SharedMailboxExpiryDays days in CustomAttribute15" }
if ($SupervisorUpn) {
  if ($GrantSupervisorFullAccess) { Add-Plan 'Mailbox' "Grant FullAccess to $SupervisorUpn" }
  if ($GrantSupervisorSendAs)     { Add-Plan 'Mailbox' "Grant SendAs to $SupervisorUpn"     }
}

$staticDLs = @($Before.EXO.DLs | Where-Object { -not $_.IsDynamic })
if ($RemoveFromDistributionLists -and $staticDLs.Count -gt 0) {
  Add-Plan 'EXO/DLs' "Remove user from $($staticDLs.Count) static distribution / mail-enabled security groups"
}

$graphStatic = @($Before.Graph.Groups | Where-Object { -not $_.IsDynamic })
if ($RemoveFromGroups -and $graphStatic.Count -gt 0) {
  Add-Plan 'Graph/Groups' "Remove user from $($graphStatic.Count) static M365/Security groups"
}

if ($RemoveMailboxDelegations -and $Before.EXO.Delegations.Count -gt 0) {
  Add-Plan 'Mailbox' "Remove $($Before.EXO.Delegations.Count) mailbox delegation entries"
}

if ($RemoveLicenses -and $Before.Licenses.Count -gt 0) {
  Add-Plan 'Licensing' "Remove all assigned licenses"
}

if ($DisableEntraSignIn) { Add-Plan 'Entra' "Block sign-in and revoke refresh tokens" }

if ($HaveAD -and ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn)) {
  Add-Plan 'AD' "Apply AD actions (Disable=$DisableAD; UpdateDesc=$UpdateAdDescription; MoveToOU='$DisabledOuDn')"
}

# Write plan
$planPath = Join-Path $OutputFolder 'Plan-WhatWeWillDo.md'
"Offboarding plan for $($User.DisplayName) <$($User.UserPrincipalName)>" | Out-File $planPath -Encoding utf8
"Ticket: $TicketNumber"            | Out-File $planPath -Append -Encoding utf8
"Preview mode: $Preview"           | Out-File $planPath -Append -Encoding utf8
""                                 | Out-File $planPath -Append -Encoding utf8
foreach ($p in $Plan) { "- [$($p.Area)] $($p.Action)" | Out-File $planPath -Append -Encoding utf8 }

# ---------- APPLY ----------
if ($Preview) {
  Step "Preview mode — not making any changes. Use -Apply to enforce."
} else {
  Step "Applying changes"

  # Mailbox conversion & expiry stamping
  if ($willConvert) {
    try {
      Act "Converting mailbox to Shared..."
      Set-Mailbox -Identity $UserUpn -Type Shared -ErrorAction Stop
      Did "Converted to Shared"
    } catch {
      Skip "Mailbox conversion failed: $_"
    }
  }

  # Stamp expiry in CustomAttribute15
  if ($mbx) {
    try {
      $expiry = (Get-Date).AddDays([int]$SharedMailboxExpiryDays)
      $marker = ("Expires: {0:yyyy-MM-dd} ({1}d)" -f $expiry, $SharedMailboxExpiryDays)
      Act "Stamping CustomAttribute15 with '$marker'"
      Set-Mailbox -Identity $UserUpn -CustomAttribute15 $marker -ErrorAction Stop
      Did "Stamped mailbox CustomAttribute15"
    } catch { Skip "Failed to stamp CustomAttribute15: $_" }
  }

  # Supervisor rights
  if ($SupervisorUpn -and $mbx) {
    if ($GrantSupervisorFullAccess) {
      try {
        Act "Grant FullAccess to $SupervisorUpn"
        Add-MailboxPermission -Identity $UserUpn -User $SupervisorUpn -AccessRights FullAccess -AutoMapping:$true -Confirm:$false
        Did "Granted FullAccess"
      } catch { Skip "Grant FullAccess failed: $_" }
    }
    if ($GrantSupervisorSendAs) {
      try {
        Act "Grant SendAs to $SupervisorUpn"
        Add-RecipientPermission -Identity $UserUpn -Trustee $SupervisorUpn -AccessRights SendAs -Confirm:$false
        Did "Granted SendAs"
      } catch { Skip "Grant SendAs failed: $_" }
    }
  }

  # Remove from EXO DLs (static only)
  if ($RemoveFromDistributionLists -and $staticDLs.Count -gt 0) {
    foreach ($g in $staticDLs) {
      try {
        Act "Remove from DL: $($g.DisplayName) <$($g.PrimarySmtp)>"
        Remove-DistributionGroupMember -Identity $g.PrimarySmtp -Member $UserUpn -BypassSecurityGroupManagerCheck -Confirm:$false
        Did "Removed from EXO DL: $($g.DisplayName)"
      } catch { Skip "DL removal failed for $($g.DisplayName): $_" }
    }
  }

  # Remove mailbox delegations
  if ($RemoveMailboxDelegations -and $Before.EXO.Delegations.Count -gt 0) {
    foreach ($d in $Before.EXO.Delegations) {
      try {
        switch ($d.Right) {
          'FullAccess' {
            Remove-MailboxPermission -Identity $d.Mailbox -User $UserUpn -AccessRights FullAccess -Confirm:$false
          }
          'SendAs' {
            Remove-RecipientPermission -Identity $d.Mailbox -Trustee $UserUpn -AccessRights SendAs -Confirm:$false
          }
          'SendOnBehalf' {
            Set-Mailbox -Identity $d.Mailbox -GrantSendOnBehalfTo @{ Remove = $UserUpn }
          }
        }
        Did "Removed mailbox delegation: $($d.Right) on $($d.Mailbox)"
      } catch { Skip "Delegation removal failed for $($d.Mailbox) [$($d.Right)]: $_" }
    }
  }


  # Backup owner on groups where the user is the sole owner
  if ($BackupOwnerUpn) {
    try {
      $backup = Resolve-GraphUser -Identity $BackupOwnerUpn
      $soleOwner = $Before.Graph.Owns | Where-Object { $_.OwnersCount -le 1 }
      foreach ($o in $soleOwner) {
        try {
          Act "Adding backup owner '$($backup.UserPrincipalName)' to group: $($o.DisplayName)"
          Add-MgGroupOwnerByRef -GroupId $o.GroupId -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($backup.Id)" } | Out-Null
          Did "Added backup owner to $($o.DisplayName)"
        } catch { Skip "Failed to add backup owner to $($o.DisplayName): $_" }
      }
    } catch { Skip "Backup owner '$BackupOwnerUpn' not resolved in Graph: $_" }
  }
  # Remove from static Graph groups
  if ($RemoveFromGroups -and $graphStatic.Count -gt 0) {
    foreach ($g in $graphStatic) {
      try {
        Act "Remove from Graph group: $($g.DisplayName)"
        # Remove membership by ref
        Remove-MgGroupMemberByRef -GroupId $g.GroupId -DirectoryObjectId $User.Id -Confirm:$false
        Did "Removed from group: $($g.DisplayName)"
      } catch { Skip "Graph group removal failed for $($g.DisplayName): $_" }
    }
  }

  # Licenses
  if ($RemoveLicenses -and $Before.Licenses.Count -gt 0) {
    try {
      $toRemove = @($Before.Licenses | Select-Object -ExpandProperty SkuId)
      Act "Removing licenses: $(@($Before.Licenses | ForEach-Object {$_.SkuPart}) -join ', ')"
      Set-MgUserLicense -UserId $User.Id -RemoveLicenses $toRemove -AddLicenses @() | Out-Null
      Did "Removed all licenses"
    } catch { Skip "License removal failed: $_" }
  }

  # Entra sign-in
  if ($DisableEntraSignIn) {
    try {
      Act "Blocking sign-in (accountEnabled=false) and revoking sessions"
      Update-MgUser -UserId $User.Id -AccountEnabled:$false | Out-Null
      Revoke-MgUserSignInSession -UserId $User.Id | Out-Null
      Did "Blocked sign-in & revoked sessions"
    } catch { Skip "Failed to block sign-in: $_" }
  }

  # AD (optional - skipped by default)
  if ($HaveAD -and ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn)) {
    try {
      $sam = ($User.UserPrincipalName -split '@')[0]
      $adUser = Get-ADUser -Identity $sam -Properties Enabled,Description,DistinguishedName -ErrorAction Stop
      if ($DisableAD -and $adUser.Enabled) {
        Act "Disabling AD account"
        Disable-ADAccount -Identity $adUser.SamAccountName
        Did "AD account disabled"
      }
      if ($UpdateAdDescription) {
        $desc = "Offboarded $(Get-Date -Format 'yyyy-MM-dd'); Ticket $TicketNumber"
        Act "Updating AD description to '$desc'"
        Set-ADUser -Identity $adUser.SamAccountName -Description $desc
        Did "AD description updated"
      }
      if ($DisabledOuDn) {
        Act "Moving user to Disabled OU: $DisabledOuDn"
        Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $DisabledOuDn
        Did "Moved AD object to '$DisabledOuDn'"
      }
    } catch { Skip "AD actions failed: $_" }
  }
}

# ---------- AFTER snapshot ----------
Step "Snapshot AFTER"
$After = [ordered]@{
  EXO      = [ordered]@{}
  Graph    = [ordered]@{}
  Licenses = @()
  AD       = @()
}

try {
  $mbx2 = Get-Mailbox -Identity $UserUpn -ErrorAction SilentlyContinue
  if ($mbx2) {
    $After.EXO.Mailbox = @{
      PrimarySmtp          = $mbx2.PrimarySmtpAddress
      RecipientTypeDetails = $mbx2.RecipientTypeDetails
      CustomAttribute15    = $mbx2.CustomAttribute15
    }
  }
} catch { }

$After.EXO.DLs         = Snapshot-EXO-DLs        -UserSmtp $UserUpn
$After.EXO.Delegations = Snapshot-EXO-Delegations -UserSmtp $UserUpn
$After.Graph.Groups    = Snapshot-GraphGroups     -UserId   $User.Id
$After.Licenses        = Snapshot-Licenses        -UserId   $User.Id
if ($HaveAD) {
  $sam = ($User.UserPrincipalName -split '@')[0]
  $After.AD = Snapshot-ADGroups -SamAccountName $sam
}

$After.EXO.DLs         | Export-Csv -Path (Join-Path $OutputFolder 'After-EXO-DLs.csv') -NoTypeInformation -Encoding UTF8
$After.EXO.Delegations | Export-Csv -Path (Join-Path $OutputFolder 'After-EXO-Delegations.csv') -NoTypeInformation -Encoding UTF8
$After.Graph.Groups    | Export-Csv -Path (Join-Path $OutputFolder 'After-Graph-Groups.csv') -NoTypeInformation -Encoding UTF8
$After.Licenses        | Export-Csv -Path (Join-Path $OutputFolder 'After-Licenses.csv') -NoTypeInformation -Encoding UTF8
if ($After.AD) { $After.AD | Export-Csv -Path (Join-Path $OutputFolder 'After-AD.csv') -NoTypeInformation -Encoding UTF8 }

# ---------- ServiceNow work notes ----------
function Summ($label,$before,$after) {
  $b = ($before | Measure-Object).Count
  $a = ($after  | Measure-Object).Count
  return "${label}: $b → $a"
}

$notesPath = Join-Path $OutputFolder ('ServiceNow-WorkNotes.txt')

$b_staticDL = $Before.EXO.DLs  | Where-Object { -not $_.IsDynamic }
$a_staticDL = $After.EXO.DLs   | Where-Object { -not $_.IsDynamic }
$b_dynDL    = $Before.EXO.DLs  | Where-Object { $_.IsDynamic }
$a_dynDL    = $After.EXO.DLs   | Where-Object { $_.IsDynamic }

$b_graph    = $Before.Graph.Groups
$a_graph    = $After.Graph.Groups
$a_graphDyn = ($a_graph | Where-Object { $_.IsDynamic }).Count

$b_deleg    = $Before.EXO.Delegations
$a_deleg    = $After.EXO.Delegations

$b_lic      = $Before.Licenses
$a_lic      = $After.Licenses

$adBefore = $Before.AD
$adAfter  = $After.AD

$expiryDatePreview = (Get-Date).AddDays([int]$SharedMailboxExpiryDays).ToString('yyyy-MM-dd')

$workNotes = @"
Offboarding — $($User.DisplayName) <$($User.UserPrincipalName)>
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Ticket: $TicketNumber
Analyst: $env:USERNAME
Mode: $(if($Preview){"Preview (no changes)"}else{"Applied"})

Summary at a glance
- $(Summ "EXO DLs (static)" $b_staticDL $a_staticDL)  | dynamic listed: $a_graphDyn
- $(Summ "Graph groups (all)" $b_graph $a_graph)
- $(Summ "Mailbox delegations" $b_deleg $a_deleg)
- $(Summ "Assigned licenses" $b_lic $a_lic)
- $(if($HaveAD){"AD snapshot written"}else{"AD not executed"})

Mailbox
- Found mailbox: $($Before.EXO.Mailbox.RecipientTypeDetails ?? 'None')
- $( if ($willConvert) { "Converted to Shared (or already Shared). Expiry marker: $expiryDatePreview" } else { "No mailbox conversion requested" } )
$( if ($SupervisorUpn) {
    $rights = @(); if ($GrantSupervisorFullAccess) { $rights += 'FullAccess' } ; if ($GrantSupervisorSendAs) { $rights += 'SendAs' }
    "Supervisor access: " + ($rights -join ' & ') + " for $SupervisorUpn"
} )

Groups & DLs
- We do not remove dynamic membership. It is shown for visibility only.
- Removed from static EXO DLs: $(if($RemoveFromDistributionLists){"Yes (see 'After-EXO-DLs.csv')"}else{"No"})
- Removed from static Graph groups: $(if($RemoveFromGroups){"Yes (see 'After-Graph-Groups.csv')"}else{"No"})`n$(if($BackupOwnerUpn){"- Added backup owner '$BackupOwnerUpn' where user was sole owner"})

Licenses & Sign‑in
- Licenses removed: $(if($RemoveLicenses){"Yes"}else{"No"})
- Entra sign‑in blocked: $(if($DisableEntraSignIn){"Yes"}else{"No"})

Artifacts
- Before snapshots: $(Join-Path $OutputFolder 'Before-*')
- After snapshots:  $(Join-Path $OutputFolder 'After-*')
- Plan:             $(Join-Path $OutputFolder 'Plan-WhatWeWillDo.md')
- Transcript:       $TranscriptPath
"@

$workNotes | Out-File $notesPath -Encoding utf8

# ---------- disconnect & wrap ----------
try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
try { Disconnect-MgGraph | Out-Null } catch {}
Stop-Transcript | Out-Null

Write-Host "`nDone. Preview: $Preview  Evidence folder: $OutputFolder" -ForegroundColor Cyan
Write-Host "ServiceNow notes file: $notesPath" -ForegroundColor Cyan



































# =================================================================================================




































