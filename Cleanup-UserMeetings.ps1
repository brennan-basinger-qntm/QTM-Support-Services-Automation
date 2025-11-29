



<#
.SYNOPSIS
  Cancel future meetings organized by a (former) user, including room bookings.

.DESCRIPTION
  - Connects to Exchange Online and Microsoft Graph
  - Runs Remove-CalendarEvents for the specified mailbox
  - Optionally runs in Preview mode first
  - Disconnects from both services at the end

.PARAMETER UserPrincipalName
  UPN / primary SMTP of the (ex) user mailbox or shared mailbox.

.PARAMETER QueryWindowInDays
  How many days into the future to look for meetings (default: 365).

.PARAMETER PreviewOnly
  If set, only shows what would be removed; no changes are made.

.EXAMPLE
  .\Cleanup-UserMeetings.ps1 -UserPrincipalName "first.last@quantinuum.com"

.EXAMPLE
  .\Cleanup-UserMeetings.ps1 -UserPrincipalName "first.last@quantinuum.com" -QueryWindowInDays 180 -PreviewOnly
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [int]$QueryWindowInDays = 365,

    [switch]$PreviewOnly
)

Write-Host "=== Calendar cleanup for $UserPrincipalName ===" -ForegroundColor Cyan
Write-Host "Query window (days): $QueryWindowInDays" -ForegroundColor Cyan
if ($PreviewOnly) {
    Write-Host "Mode: PREVIEW ONLY (no changes will be made)" -ForegroundColor Yellow
} else {
    Write-Host "Mode: LIVE (changes WILL be made)" -ForegroundColor Red
}

try {
    # ---------------------------------------------
    # 1. Connect to Exchange Online
    # ---------------------------------------------
    Write-Host "`n[1/4] Connecting to Exchange Online..." -ForegroundColor Cyan

    # If already using Modern auth with Exchange Online module v3+
    Connect-ExchangeOnline -ShowBanner:$false

    # ---------------------------------------------
    # 2. Connect to Microsoft Graph (optional now, useful if we expand expand later)
    # ---------------------------------------------
    Write-Host "[2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan

    # Might need more scopes later for additional Graph calls (e.g. to query room mailboxes, users, etc.)
    $graphScopes = @(
        "User.Read.All",
        "Directory.Read.All"
        # Add more scopes later for calendar/room stuff via Graph
    )

    try {
        Connect-MgGraph -Scopes $graphScopes | Out-Null
        Write-Host "Connected to Microsoft Graph with scopes: $($graphScopes -join ', ')" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        Write-Warning "Continuing with Exchange Online only (Remove-CalendarEvents does NOT require Graph)."
    }

    # ---------------------------------------------
    # 3. Run Remove-CalendarEvents
    # ---------------------------------------------
    Write-Host "`n[3/4] Running Remove-CalendarEvents for $UserPrincipalName..." -ForegroundColor Cyan

    $params = @{
        Identity             = $UserPrincipalName
        CancelOrganizedMeetings = $true
        QueryWindowInDays    = $QueryWindowInDays
    }

    if ($PreviewOnly) {
        $params["PreviewOnly"] = $true
    }

    $result = Remove-CalendarEvents @params

    Write-Host "`nRemove-CalendarEvents completed." -ForegroundColor Green

    if ($PreviewOnly) {
        Write-Host "Preview mode: No meetings were actually cancelled."
        Write-Host "Review the output above, then rerun WITHOUT -PreviewOnly to commit."
    } else {
        Write-Host "Live mode: Future meetings organized by $UserPrincipalName within the next $QueryWindowInDays days have been cancelled."
    }

} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    # ---------------------------------------------
    # 4. Disconnect from services
    # ---------------------------------------------
    Write-Host "`n[4/4] Disconnecting from services..." -ForegroundColor Cyan

    try {
        # Exchange Online disconnect
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Warning "Problem disconnecting from Exchange Online: $($_.Exception.Message)"
    }

    try {
        # Microsoft Graph disconnect
        if (Get-Module -Name Microsoft.Graph -ListAvailable -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Problem disconnecting from Microsoft Graph: $($_.Exception.Message)"
    }

    Write-Host "`nAll done." -ForegroundColor Cyan
}

