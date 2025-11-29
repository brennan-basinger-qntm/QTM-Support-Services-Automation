<#
.SYNOPSIS
  Cancel future meetings organized by a (former) user, including room bookings,
  and verify from the script that they’re gone.

.DESCRIPTION
  - Connects to Exchange Online and Microsoft Graph
  - Runs Remove-CalendarEvents for the specified mailbox
  - In PREVIEW mode: shows which meetings *would* be cancelled
  - In LIVE mode: cancels meetings, then runs a Preview scan again to verify
  - Disconnects cleanly from both services

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

# Use “now” as explicit start date
$queryStart = Get-Date
Write-Host ("Query start date: {0}" -f $queryStart) -ForegroundColor Cyan

try {
    # ---------------------------------------------
    # 1. Connect to Exchange Online
    # ---------------------------------------------
    Write-Host "`n[1/4] Connecting to Exchange Online..." -ForegroundColor Cyan

    Connect-ExchangeOnline -ShowBanner:$false

    # ---------------------------------------------
    # 2. Connect to Microsoft Graph (optional / cosmetic)
    # ---------------------------------------------
    Write-Host "[2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan

    # Scopes kept minimal; we’re not actually using Graph for anything yet.
    $graphScopes = @(
        "User.Read.All",
        "Directory.Read.All"
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
        Identity               = $UserPrincipalName
        CancelOrganizedMeetings = $true
        QueryStartDate         = $queryStart
        QueryWindowInDays      = $QueryWindowInDays
    }

    if ($PreviewOnly) {
        $params["PreviewOnly"] = $true
    }

    # Force verbose output for this section so we SEE the meetings
    $oldVerbosePreference = $VerbosePreference
    $VerbosePreference = 'Continue'

    try {
        Write-Host "`n--- PRIMARY RUN (PreviewOnly = $PreviewOnly) ---`n" -ForegroundColor DarkCyan
        $result = Remove-CalendarEvents @params -Verbose

        # Try to give a count as well
        if ($null -ne $result) {
            $count = ($result | Measure-Object).Count
            Write-Host "`nRemove-CalendarEvents returned $count item(s)." -ForegroundColor Green
        } else {
            Write-Host "`nRemove-CalendarEvents did not return an object; rely on the verbose lines above for details." -ForegroundColor Yellow
        }

        if ($PreviewOnly) {
            Write-Host "`nPreview mode: No meetings were actually cancelled." -ForegroundColor Yellow
            Write-Host "Review the verbose output above to see exactly which meetings WOULD be removed." -ForegroundColor Yellow
        } else {
            Write-Host "`nLive mode: Future meetings organized by $UserPrincipalName within the next $QueryWindowInDays days have been processed." -ForegroundColor Green

            # ---------------------------------------------
            # 3b. Verification run (PreviewOnly AFTER deletion)
            # ---------------------------------------------
            Write-Host "`n[Verification] Running a PreviewOnly scan AFTER deletion..." -ForegroundColor Cyan

            $verifyParams = @{
                Identity               = $UserPrincipalName
                CancelOrganizedMeetings = $true
                QueryStartDate         = $queryStart
                QueryWindowInDays      = $QueryWindowInDays
                PreviewOnly            = $true
            }

            Write-Host "`n--- VERIFICATION RUN (PreviewOnly = True) ---`n" -ForegroundColor DarkCyan
            $verifyResult = Remove-CalendarEvents @verifyParams -Verbose

            if ($null -ne $verifyResult) {
                $verifyCount = ($verifyResult | Measure-Object).Count
                Write-Host "`nVerification: Remove-CalendarEvents still sees $verifyCount matching item(s) in this window." -ForegroundColor Yellow
                Write-Host "Check the verbose output above; if there are still meetings listed, they were NOT removed." -ForegroundColor Yellow
            } else {
                Write-Host "`nVerification: Remove-CalendarEvents reported no remaining meetings in this window." -ForegroundColor Green
                Write-Host "If no meetings were listed in the verification run’s verbose output, the target series is gone." -ForegroundColor Green
            }
        }
    } finally {
        # Restore original verbose preference
        $VerbosePreference = $oldVerbosePreference
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
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Warning "Problem disconnecting from Exchange Online: $($_.Exception.Message)"
    }

    try {
        if (Get-Module -Name Microsoft.Graph -ListAvailable -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Problem disconnecting from Microsoft Graph: $($_.Exception.Message)"
    }

    Write-Host "`nAll done." -ForegroundColor Cyan
}
