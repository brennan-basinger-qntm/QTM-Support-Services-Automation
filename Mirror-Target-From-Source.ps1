





# ==================================================================================================
# READ ME PLEASE

# Mirror-Source-To-Target.ps1
# --------------------------------------------------------------------------------------------------
# PURPOSE
#   • Mirror static Distribution Lists & mail-enabled security groups (Exchange Online)
#   • Mirror Shared Mailbox permissions: FullAccess, SendAs, SendOnBehalf (Exchange Online)
#   • Mirror Microsoft 365 Groups / Teams membership & ownership (Microsoft Graph)
#
# SAFETY
#   • PREVIEW mode is ON by default (set $Preview = $false to actually apply changes)
#   • Every major section is in try/catch so one failure doesn’t kill the run
#   • A transcript is written to Desktop for audit/rollback
#
# ONE-TIME REQUIREMENTS
#   • Windows PowerShell 5.1 or PowerShell 7.4+
#   • Install Modules (admin recommended): ExchangeOnlineManagement, Microsoft.Graph
#   • Graph delegated scopes (consent on first run): GroupMember.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All
# ==================================================================================================




# ----------------------------- USER SETTINGS (EDIT THESE) -----------------------------------------
# [Block] Define source/target and behavior flags.

# Define the SOURCE user whose access/memberships you want to COPY FROM.
$Source  = 'benjamin.spaun@quantinuum.com'

# Define the TARGET user who should RECEIVE the same access/memberships.
$Target  = 'thomas.wilkason@quantinuum.com'

# Start in PREVIEW mode (no changes). Set to $false to APPLY.
$Preview = $true

# Optional: Hint Graph which tenant to use (tenant GUID or verified domain like 'quantinuum.com').
# Leave $null to let the sign-in picker choose.
# We'll make it easy on the program and give it the tenant GUID 
# Tenant ID can be found Entra -> Overview -> Properties

# This should work without explicitly defining the tenant GUID 
# Script would still work if we set it to $Null
$TenantHint = '94c4857e-1130-4ab8-8eac-069b40c9db20'
# --------------------------------------------------------------------------------------------------



# ----------------------------- TRANSCRIPT (LOGGING) -----------------------------------------------
# Build a Desktop path for the transcript with timestamp.
$TranscriptPath = Join-Path $env:USERPROFILE ("Desktop\Mirror-Run-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt")

# Begin recording console output and errors to that transcript.
Start-Transcript -Path $TranscriptPath -Append
# --------------------------------------------------------------------------------------------------


# ----------------------------- CONSOLE HELPERS ----------------------------------------------------
# Small helpers for consistent console output - I've got the transcript currently capturing text only.
function Write-Info($msg) { Write-Host $msg }
function Write-Step($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-Act ($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Skip($msg) { Write-Warning $msg }
# --------------------------------------------------------------------------------------------------


# ----------------------------- MODULE INSTALL/IMPORT ----------------------------------------------
# Ensure a module exists; install (AllUsers) and import. If AllUsers fails, fall back to CurrentUser.
function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)

  # If module not found locally, try to install for all users.
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Step "Installing module: $Name (AllUsers)..."
    try {
      Install-Module $Name -Scope AllUsers -Force -ErrorAction Stop
    } catch {
      # [Line] If AllUsers install fails (no admin), try CurrentUser.
      Write-Step "AllUsers install failed; installing $Name for CurrentUser..."
      Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
  }

  # Import the module into the session.
  Import-Module $Name -ErrorAction Stop
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- GRAPH PREFLIGHT (CONNECT + VERIFY) --------------------------------
# [Block] Connect to Microsoft Graph with exact scopes, verify tenant/account/scopes, and select v1.0.
function Ensure-Graph {
  param(
    [string]$TenantIdOrDomain = $null,
    [string[]]$RequiredScopes = @('GroupMember.ReadWrite.All','Group.ReadWrite.All','Directory.Read.All')
  )

  # [Line] Ensure Graph SDK is available.
  Ensure-Module -Name Microsoft.Graph

  # [Line] Try to reuse existing context if present.
  $ctx = Get-MgContext -ErrorAction SilentlyContinue

  # [Block] Decide if we need to (re)connect: missing context, wrong tenant, or missing scopes.
  $needConnect = $true
  if ($ctx) {
    $tenantMismatch = $false
    if ($TenantIdOrDomain) {
      # [Line] If TenantHint is set, treat a different tenant as mismatch.
      $tenantMismatch = ($ctx.TenantId -ne $TenantIdOrDomain -and $ctx.TenantId -ne $null -and $TenantIdOrDomain -ne $null)
    }
    $missingScopes = @()
    if ($ctx.Scopes) { $missingScopes = $RequiredScopes | Where-Object { $_ -notin $ctx.Scopes } }
    $needConnect = $tenantMismatch -or ($missingScopes.Count -gt 0)
  }

  # [Block] Connect if needed with the required scopes and chosen tenant.
  if ($needConnect) {
    if ($TenantIdOrDomain) {
      Write-Step "Connecting to Graph tenant '$TenantIdOrDomain' with scopes: $($RequiredScopes -join ', ')"
      Connect-MgGraph -Scopes $RequiredScopes -TenantId $TenantIdOrDomain
    } else {
      Write-Step "Connecting to Graph (no tenant hint) with scopes: $($RequiredScopes -join ', ')"
      Connect-MgGraph -Scopes $RequiredScopes
    }
    # [Line] Choose the stable v1.0 profile.
    Select-MgProfile -Name 'v1.0'
    # [Line] Refresh context post-connect.
    $ctx = Get-MgContext
  } else {
    # [Line] Ensure we’re on v1.0 even if we reused context.
    Select-MgProfile -Name 'v1.0'
  }

  # [Block] Validate granted scopes.
  $granted = @($ctx.Scopes)
  $missing = $RequiredScopes | Where-Object { $_ -notin $granted }
  if ($missing.Count -gt 0) {
    throw "Missing Graph scopes: $($missing -join ', '). Re-run Connect-MgGraph with the full set."
  }

  # [Line] Print who/where we are and what scopes we have.
  Write-Host "Graph account  : $($ctx.Account)" -ForegroundColor Cyan
  Write-Host "Graph tenant   : $($ctx.TenantId)" -ForegroundColor Cyan
  Write-Host "Granted scopes : $($granted -join ', ')" -ForegroundColor Cyan
}

# [Line] Call Graph preflight now (early fail if something is wrong).
Ensure-Graph -TenantIdOrDomain $TenantHint
# --------------------------------------------------------------------------------------------------


# ----------------------------- EXCHANGE ONLINE CONNECT --------------------------------------------
# [Block] Load EXO v3 (REST-backed) and connect; no WinRM Basic needed.
try {
  Ensure-Module -Name ExchangeOnlineManagement
  Connect-ExchangeOnline -ShowBanner:$false
} catch {
  Write-Skip "Exchange Online connection failed: $_"
  throw
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- RESOLVE SOURCE/TARGET IN EXO ---------------------------------------
# [Block] Resolve SMTPs to recipient objects to catch typos/aliases early.
try {
  $src = Get-Recipient -Identity $Source -ErrorAction Stop
  $dst = Get-Recipient -Identity $Target -ErrorAction Stop
} catch {
  Write-Skip "Could not resolve Source or Target in Exchange: $_"
  throw
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- DISTRIBUTION LISTS (STATIC ONLY) -----------------------------------
# [Block] DL mirror: only static DLs / mail-enabled security groups, no dynamic DLs.
Write-Step "== DISTRIBUTION LISTS (static only; dynamic DLs are skipped) =="
try {
  # [Line] Get all non-dynamic DLs.
  $groups = Get-DistributionGroup -ResultSize Unlimited |
            Where-Object { $_.RecipientTypeDetails -ne 'DynamicDistributionGroup' }

  # [Block] Loop each DL and mirror membership if Source is a member.
  foreach ($g in $groups) {
    try {
      # [Line] Grab all current members of the DL.
      $members = Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited

      # [Line] Check if Source belongs to this DL.
      $srcIn   = $members | Where-Object { $_.PrimarySmtpAddress -ieq $Source }
      if ($srcIn) {
        # [Line] Check if Target is already in; avoid duplicates.
        $tgtIn = $members | Where-Object { $_.PrimarySmtpAddress -ieq $Target }

        # [Block] Add Target when missing (apply only if Preview = $false).
        if (-not $tgtIn) {
          Write-Act "ADD $Target to DL: $($g.DisplayName)"
          if (-not $Preview) {
            Add-DistributionGroupMember -Identity $g.Identity -Member $Target -BypassSecurityGroupManagerCheck -ErrorAction Stop
          }
        } else {
          Write-Info "Already in: $($g.DisplayName)"
        }
      }
    } catch {
      # [Line] Keep going even if this DL hits an error.
      Write-Skip "DL $($g.DisplayName): $_"
    }
  }
} catch {
  Write-Skip "DL enumeration failed: $_"
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- SHARED MAILBOX PERMISSIONS -----------------------------------------
# [Block] Mirror FullAccess, SendAs, and SendOnBehalf for shared mailboxes.
Write-Step "`n== SHARED MAILBOX PERMISSIONS (FullAccess / SendAs / SendOnBehalf) =="
try {
  # [Line] Get all shared mailboxes.
  $shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

  # [Block] Loop each shared mailbox and mirror the three permission types.
  foreach ($mbx in $shared) {
    try {
      # ----- FullAccess -----
      # [Line] Find explicit FullAccess grants for Source (non-inherited, not deny).
      $faSrc = Get-MailboxPermission -Identity $mbx.Identity |
               Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.AccessRights -contains 'FullAccess' -and $_.User.ToString() -ieq $Source }

      # [Block] If Source has FullAccess, ensure Target does too.
      if ($faSrc) {
        $faTgt = Get-MailboxPermission -Identity $mbx.Identity |
                 Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.AccessRights -contains 'FullAccess' -and $_.User.ToString() -ieq $Target }
        if (-not $faTgt) {
          Write-Act "GRANT FullAccess on $($mbx.PrimarySmtpAddress)"
          if (-not $Preview) {
            Add-MailboxPermission -Identity $mbx.Identity -User $Target -AccessRights FullAccess -Confirm:$false
          }
        }
      }

      # ----- SendAs -----
      # [Line] Find SendAs for Source at the recipient level.
      $saSrc = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
               Where-Object { $_.Trustee -ieq $Source -and $_.AccessRights -contains 'SendAs' }

      # [Block] If Source has SendAs, ensure Target does too.
      if ($saSrc) {
        $saTgt = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                 Where-Object { $_.Trustee -ieq $Target -and $_.AccessRights -contains 'SendAs' }
        if (-not $saTgt) {
          Write-Act "GRANT SendAs on $($mbx.PrimarySmtpAddress)"
          if (-not $Preview) {
            Add-RecipientPermission -Identity $mbx.Identity -Trustee $Target -AccessRights SendAs -Confirm:$false
          }
        }
      }

      # ----- SendOnBehalf -----
      # [Line] Get current SendOnBehalf delegates.
      $sobList = (Get-Mailbox -Identity $mbx.Identity).GrantSendOnBehalfTo

      # [Block] If there are delegates, compare using SMTPs and add Target if Source has it.
      if ($sobList) {
        $sobSmtp = $sobList | ForEach-Object { (Get-Recipient $_).PrimarySmtpAddress }
        if ($sobSmtp -contains $Source -and -not ($sobSmtp -contains $Target)) {
          Write-Act "GRANT SendOnBehalf on $($mbx.PrimarySmtpAddress)"
          if (-not $Preview) {
            Set-Mailbox -Identity $mbx.Identity -GrantSendOnBehalfTo @{ Add = $Target }
          }
        }
      }

    } catch {
      # [Line] Keep going even if this mailbox hits an error.
      Write-Skip "Mailbox $($mbx.PrimarySmtpAddress): $_"
    }
  }
} catch {
  Write-Skip "Shared mailbox enumeration failed: $_"
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- GRAPH UTIL: USER RESOLUTION ----------------------------------------
# [Block] Helper to resolve a user in Graph by UPN/mail; returns a user with Id/Mail/UPN.
function Resolve-MgUser {
  param([Parameter(Mandatory)][string]$Identity)

  # [Line] Try direct lookup by UPN/objectId.
  $u = $null
  try { $u = Get-MgUser -UserId $Identity -ErrorAction Stop } catch { }

  # [Line] If that failed, filter by mail or UPN and take the first match.
  if (-not $u) {
    $escaped = $Identity.Replace("'","''")
    $u = Get-MgUser -All -Filter "mail eq '$escaped' or userPrincipalName eq '$escaped'" | Select-Object -First 1
  }

  # [Line] If still nothing, throw a clear error.
  if (-not $u) { throw "User not found in Microsoft Graph: $Identity" }
  return $u
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- MICROSOFT GRAPH: M365 GROUPS/TEAMS ---------------------------------
# [Block] Mirror membership & ownership for M365 Groups (Unified groups) via Graph.
Write-Step "`n== M365 GROUPS (Unified Groups / Teams) via Microsoft Graph =="
try {
  # [Line] Resolve Source and Target in Graph to get stable object IDs.
  $srcUser = Resolve-MgUser -Identity $Source
  $tgtUser = Resolve-MgUser -Identity $Target

  # [Line] Get all Microsoft 365 groups (Unified = M365/Teams).
  $mgGroups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')"

  # [Block] Loop each group to mirror where Source participates.
  foreach ($g in $mgGroups) {
    try {
      # ----- MEMBERS -----
      # [Line] Read members (DirectoryObject records) and collect their IDs.
      $members   = Get-MgGroupMember -GroupId $g.Id -All
      $memberIds = @($members | ForEach-Object { $_.Id })

      # [Block] If Source is member and Target isn’t, add Target as member.
      $srcIsMember = $memberIds -contains $srcUser.Id
      if ($srcIsMember) {
        $tgtIsMember = $memberIds -contains $tgtUser.Id
        if (-not $tgtIsMember) {
          Write-Act "ADD $($tgtUser.Mail ?? $Target) as MEMBER of: $($g.DisplayName)"
          if (-not $Preview) {
            # [Line] Add by reference (Graph owner/member adds use @odata.id references).
            New-MgGroupMemberByRef -GroupId $g.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($tgtUser.Id)" }
          }
        } else {
          Write-Info "Already a MEMBER: $($g.DisplayName)"
        }
      }

      # ----- OWNERS -----
      # [Line] Read owners and collect IDs.
      $owners   = Get-MgGroupOwner -GroupId $g.Id -All
      $ownerIds = @($owners | ForEach-Object { $_.Id })

      # [Block] If Source is owner and Target isn’t, add Target as owner.
      $srcIsOwner = $ownerIds -contains $srcUser.Id
      if ($srcIsOwner) {
        $tgtIsOwner = $ownerIds -contains $tgtUser.Id
        if (-not $tgtIsOwner) {
          Write-Act "ADD $($tgtUser.Mail ?? $Target) as OWNER of: $($g.DisplayName)"
          if (-not $Preview) {
            New-MgGroupOwnerByRef -GroupId $g.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($tgtUser.Id)" }
          }
        } else {
          Write-Info "Already an OWNER: $($g.DisplayName)"
        }

        # [Block] Ensure owners are also members (Teams expects this).
        if (-not ($memberIds -contains $tgtUser.Id)) {
          Write-Act "Ensure OWNER is also MEMBER for: $($g.DisplayName)"
          if (-not $Preview) {
            New-MgGroupMemberByRef -GroupId $g.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($tgtUser.Id)" }
          }
        }
      }

    } catch {
      # [Line] Log per-group errors and continue.
      Write-Skip "M365 Group $($g.DisplayName): $_"
    }
  }
} catch {
  Write-Skip "M365 Groups via Graph failed: $_"
}
# --------------------------------------------------------------------------------------------------


# ----------------------------- CLEANUP ------------------------------------------------------------
# [Block] Try to disconnect Graph quietly (safe even if not connected).
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }

# [Line] Close Exchange Online session without prompting.
Disconnect-ExchangeOnline -Confirm:$false

# [Line] Stop the transcript and print a summary line with preview state and log path.
Stop-Transcript
Write-Host "`nPreview mode: $Preview (set to `$false to APPLY) - Transcript: $TranscriptPath"
# --------------------------------------------------------------------------------------------------





