<#
Fast-AuditUser.ps1
Read only snapshot for ServiceNow: user, licenses, groups, and optional mailbox delegates.

Design goals:
- Fast default run
- Avoid tenant-wide enumerations unless you opt in
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $UserPrincipalName,

    # If not provided, we will resolve this in Resolve-OutFolder
    [Parameter(Mandatory = $false)]
    [string] $OutFolder,

    # Console display controls
    [int] $MaxConsoleRows = 60,

    # Microsoft Graph group collection controls
    [int] $MaxGroupsToFetch = 200,
    [switch] $FetchAllGroups,

    # Exchange Online controls
    [switch] $IncludeExchangeOnline,
    [switch] $IncludeMailboxDelegates,

    # This can be slower because it may enumerate many objects tenant-wide
    [switch] $IncludeOwnedMessagingGroups
)

$ErrorActionPreference = 'Stop'

function Resolve-OutFolder {
    param([string] $Requested)

    if (-not [string]::IsNullOrWhiteSpace($Requested)) {
        return $Requested
    }

    # Preferred: the run folder created by your PowerShell profile.
    # Your profile sets $global:ITOPS_RUN_FOLDER under C:\IT\OpsEnv\runs\<timestamp>. [1](https://quantinuum-my.sharepoint.com/personal/brennan_basinger_quantinuum_com/_layouts/15/Doc.aspx?action=edit&mobileredirect=true&wdorigin=Sharepoint&DefaultItemOpen=1&sourcedoc={61d3d376-3f72-4458-8086-2cf4979c7752}&wd=target(/SupportServices.one/)&wdpartid={3ed9082a-d023-0b0c-0886-b2ea06ec0189}{1}&wdsectionfileid={1245c07e-ac5f-4d3a-a1d7-1a1840d4544a})[2](https://quantinuum-my.sharepoint.com/personal/brennan_basinger_quantinuum_com/Documents/Microsoft%20Teams%20Chat%20Files/Offboarding_Process_Overview.html?web=1)
    if (-not [string]::IsNullOrWhiteSpace($env:ITOPS_RUN_FOLDER)) {
        return $env:ITOPS_RUN_FOLDER
    }
    if (-not [string]::IsNullOrWhiteSpace($global:ITOPS_RUN_FOLDER)) {
        return $global:ITOPS_RUN_FOLDER
    }

    # Fallback: create a local folder in the current directory so it is obvious.
    $runId = Get-Date -Format "yyyy-MM-dd_HH.mm.ss"
    return (Join-Path (Get-Location).Path ("AuditUser_" + $runId))
}

function Ensure-Folder {
    param([string] $Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Ensure-GraphConnection {
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Directory.Read.All"
    }
    try {
        Select-MgProfile -Name "v1.0" | Out-Null
    } catch { }
}

function Ensure-ExchangeConnection {
    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {

        # First attempt: disable Web Account Manager (avoids common token broker issues).
        try {
            Connect-ExchangeOnline -DisableWAM -ShowBanner:$false | Out-Null
            return
        } catch { }

        # Fallback: device flow
        Connect-ExchangeOnline -Device -ShowBanner:$false | Out-Null
    }
}

# -------------------- start --------------------
$OutFolder = Resolve-OutFolder -Requested $OutFolder
Ensure-Folder -Path $OutFolder

$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$timings = [System.Collections.Generic.List[object]]::new()

# -------------------- Graph snapshot --------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Ensure-GraphConnection

# Pull minimal user properties only (faster and cleaner)
$u = Get-MgUser -UserId $UserPrincipalName -Property "id,displayName,userPrincipalName,mail,accountEnabled,createdDateTime,userType"

# Licenses: retrieve details first, then format
$licenseDetails = @()
try {
    $licenseDetails = Get-MgUserLicenseDetail -UserId $u.Id -All -ErrorAction Stop
} catch {
    $licenseDetails = @()
}

# Licenses: keep it simple and safe
$licensesPretty = @()
foreach ($l in $licenseDetails) {

    # SkuId can be a Guid already, or an object with a Guid property depending on module versions
    $skuIdValue = $null
    if ($null -ne $l.SkuId) {
        if ($l.SkuId -is [Guid]) { $skuIdValue = $l.SkuId }
        elseif ($null -ne $l.SkuId.Guid) { $skuIdValue = $l.SkuId.Guid }
        else { $skuIdValue = [string]$l.SkuId }
    }

    $licensesPretty += [pscustomobject]@{
        LicenseName = $l.SkuPartNumber
        SkuId       = $skuIdValue
    }
}

# Groups
$groups = @()
try {
    if ($FetchAllGroups) {
        $objs = Get-MgUserMemberOf -UserId $u.Id -All
    } else {
        $objs = Get-MgUserMemberOf -UserId $u.Id -Top $MaxGroupsToFetch
    }

    $groups = $objs |
        Where-Object {
            $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' -or $_.'@odata.type' -like '*group*'
        } |
        Select-Object @{n='GroupName';e={$_.DisplayName}}, @{n='GroupId';e={$_.Id}} |
        Sort-Object GroupName
} catch {
    $groups = @()
}

$sw.Stop()
$timings.Add([pscustomobject]@{ Section = "Microsoft Graph snapshot"; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2) }) | Out-Null

# -------------------- Exchange Online snapshot (optional) --------------------
$mailbox = $null
$fullAccess = @()
$sendAs = @()
$ownedDLists = @()
$ownedM365Groups = @()

if ($IncludeExchangeOnline -or $IncludeMailboxDelegates -or $IncludeOwnedMessagingGroups) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Ensure-ExchangeConnection

    $mailbox = Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue

    if ($mailbox -and ($IncludeMailboxDelegates -or $IncludeExchangeOnline)) {

        $fullAccess = Get-MailboxPermission -Identity $mailbox.Identity |
            Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited } |
            Select-Object User,AccessRights

        $sendAs = Get-RecipientPermission -Identity $mailbox.Identity |
            Where-Object { $_.AccessRights -contains "SendAs" -and -not $_.IsInherited } |
            Select-Object Trustee,AccessRights
    }

    if ($IncludeOwnedMessagingGroups) {
        $ownedDLists = Get-DistributionGroup -ResultSize Unlimited |
            Where-Object { $_.ManagedBy -match [regex]::Escape($UserPrincipalName) } |
            Select-Object DisplayName,PrimarySmtpAddress

        $ownedM365Groups = Get-UnifiedGroup -ResultSize Unlimited |
            Where-Object { $_.Owners -contains $UserPrincipalName } |
            Select-Object DisplayName,PrimarySmtpAddress
    }

    $sw.Stop()
    $timings.Add([pscustomobject]@{ Section = "Exchange Online snapshot"; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2) }) | Out-Null
}

# -------------------- Console summary --------------------
Write-Host ""
Write-Host "User snapshot" -ForegroundColor Cyan
Write-Host ("Name:            {0}" -f $u.DisplayName)
Write-Host ("User principal:  {0}" -f $u.UserPrincipalName)
Write-Host ("Mail:            {0}" -f $u.Mail)
Write-Host ("Account enabled: {0}" -f $u.AccountEnabled)
Write-Host ("User type:       {0}" -f $u.UserType)
Write-Host ("Created:         {0}" -f $u.CreatedDateTime)

Write-Host ""
Write-Host ("Licenses:        {0}" -f @($licensesPretty).Count) -ForegroundColor Cyan
$licensesPretty | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize
if (@($licensesPretty).Count -gt $MaxConsoleRows) {
    Write-Host ("Showing first {0}. Full list is in licenses.csv." -f $MaxConsoleRows) -ForegroundColor Yellow
}

Write-Host ""
Write-Host ("Groups returned: {0}" -f @($groups).Count) -ForegroundColor Cyan
$groups | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize
if (@($groups).Count -gt $MaxConsoleRows) {
    Write-Host ("Showing first {0}. Full list is in groups.csv." -f $MaxConsoleRows) -ForegroundColor Yellow
}

if ($mailbox) {
    Write-Host ""
    Write-Host "Mailbox" -ForegroundColor Cyan
    Write-Host ("Mailbox found:   {0}" -f $mailbox.RecipientTypeDetails)

    if ($IncludeMailboxDelegates -or $IncludeExchangeOnline) {
        Write-Host ""
        Write-Host ("Full Access entries: {0}" -f @($fullAccess).Count) -ForegroundColor Cyan
        $fullAccess | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize

        Write-Host ""
        Write-Host ("Send As entries:     {0}" -f @($sendAs).Count) -ForegroundColor Cyan
        $sendAs | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize
    }

    if ($IncludeOwnedMessagingGroups) {
        Write-Host ""
        Write-Host ("Owned distribution lists: {0}" -f @($ownedDLists).Count) -ForegroundColor Cyan
        $ownedDLists | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize

        Write-Host ""
        Write-Host ("Owned Microsoft 365 Groups: {0}" -f @($ownedM365Groups).Count) -ForegroundColor Cyan
        $ownedM365Groups | Select-Object -First $MaxConsoleRows | Format-Table -AutoSize
    }
}

# -------------------- Exports --------------------
if (-not [string]::IsNullOrWhiteSpace($OutFolder)) {

    ($u | Select-Object Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,CreatedDateTime,UserType) |
        ConvertTo-Json | Out-File (Join-Path $OutFolder "user.json") -Encoding UTF8

    $licensesPretty | Export-Csv (Join-Path $OutFolder "licenses.csv") -NoTypeInformation
    $groups | Export-Csv (Join-Path $OutFolder "groups.csv") -NoTypeInformation

    if ($mailbox -and ($IncludeMailboxDelegates -or $IncludeExchangeOnline)) {
        $fullAccess | Export-Csv (Join-Path $OutFolder "mailbox_fullaccess.csv") -NoTypeInformation
        $sendAs | Export-Csv (Join-Path $OutFolder "mailbox_sendas.csv") -NoTypeInformation
    }

    if ($IncludeOwnedMessagingGroups) {
        $ownedDLists | Export-Csv (Join-Path $OutFolder "owned_distribution_lists.csv") -NoTypeInformation
        $ownedM365Groups | Export-Csv (Join-Path $OutFolder "owned_m365_groups.csv") -NoTypeInformation
    }

    $timings | Export-Csv (Join-Path $OutFolder "timings.csv") -NoTypeInformation
}

$swTotal.Stop()

Write-Host ""
Write-Host "Timings" -ForegroundColor Cyan
$timings | Format-Table -AutoSize
Write-Host ("Total seconds: {0}" -f [math]::Round($swTotal.Elapsed.TotalSeconds, 2)) -ForegroundColor Cyan

# Always show the output folder and what files were written
Write-Host ""
Write-Host ("Output folder: {0}" -f $OutFolder) -ForegroundColor Green
Write-Host "Files written:" -ForegroundColor Green
Get-ChildItem -Path $OutFolder -File |
    Select-Object Name,Length,LastWriteTime |
    Format-Table -AutoSize

# Return the path as the final output so it is easy to copy
$OutFolder

# Optional cleanup
# Disconnect-ExchangeOnline -Confirm:$false | Out-Null
# Disconnect-MgGraph | Out-Null