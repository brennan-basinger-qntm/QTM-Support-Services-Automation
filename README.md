# ==================================================================================================
# ==================================================================================================
# READ ME PLEASE
# ==================================================================================================
# ==================================================================================================


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
#   • PowerShell 7.4+ (x64)
#   • Install Modules (admin recommended): ExchangeOnlineManagement, Microsoft.Graph
#   • Graph delegated scopes (consent on first run): GroupMember.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All
# ==================================================================================================




# ----------------------------- USER SETTINGS (EDIT THESE) -----------------------------------------
# Define source/target and behavior flags

# Define the SOURCE user whose access/memberships you want to COPY FROM.
$Source  = 'jacob.underwood@quantinuum.com'

# Define the TARGET user who should RECEIVE the same access/memberships.
$Target  = 'brennan.basinger@quantinuum.com'

# Start in PREVIEW mode (no changes). Set to $false to APPLY.
$Preview = $true

# Don't touch anything below this line
# ----------------------------- USER SETTINGS (EDIT THESE) -----------------------------------------








