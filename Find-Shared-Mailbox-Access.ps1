# FindSharedMailboxAccess.ps1

# Find-Mailbox-Access.ps1
# How to run this script:
# .\Find-Mailbox-Access.ps1 -UserEmail "koko.breadmore@quantinuum.com"
# .\Find-Mailbox-Access.ps1 -UserEmail "koko.breadmore@quantinuum.com" -SharedMailboxesOnly

param(
    [Parameter(Mandatory = $true)]
    [string]$UserEmail,

    [string]$OutputFolder = "C:\Temp\MailboxAccessAudit",

    [switch]$SharedMailboxesOnly
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "== $Message ==" -ForegroundColor Cyan
}

function New-IdentitySet {
    return New-Object 'System.Collections.Generic.HashSet[string]'
}

function Add-NormalizedValue {
    param(
        [System.Collections.Generic.HashSet[string]]$Set,
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        [void]$Set.Add($Value.Trim().ToLowerInvariant())
    }
}

function Get-SafeString {
    param($Value)

    if ($null -eq $Value) { return $null }
    return [string]$Value
}

function Has-Property {
    param(
        $Object,
        [string]$Name
    )

    try {
        return ($null -ne $Object -and $Object.PSObject.Properties.Match($Name).Count -gt 0)
    }
    catch {
        return $false
    }
}

function Expand-RecipientIdentityStrings {
    param($Recipient)

    $set = New-IdentitySet

    if ($null -eq $Recipient) {
        return $set
    }

    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.Name)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.DisplayName)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.Alias)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.PrimarySmtpAddress)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.WindowsEmailAddress)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.UserPrincipalName)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.DistinguishedName)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.Guid)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.ExternalDirectoryObjectId)
    Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.Identity)

    if (Has-Property -Object $Recipient -Name 'LegacyExchangeDN') {
        Add-NormalizedValue -Set $set -Value (Get-SafeString $Recipient.LegacyExchangeDN)
    }

    return $set
}

function Merge-HashSet {
    param(
        [System.Collections.Generic.HashSet[string]]$Target,
        [System.Collections.Generic.HashSet[string]]$Source
    )

    foreach ($item in $Source) {
        [void]$Target.Add($item)
    }
}

function Resolve-IdentityStringToSet {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$IdentityString
    )

    $set = New-IdentitySet
    Add-NormalizedValue -Set $set -Value $IdentityString

    if ([string]::IsNullOrWhiteSpace($IdentityString)) {
        return $set
    }

    try {
        $recipient = Get-Recipient -Identity $IdentityString -ErrorAction Stop
        $recipientSet = Expand-RecipientIdentityStrings -Recipient $recipient
        Merge-HashSet -Target $set -Source $recipientSet
    }
    catch {
        # Keep raw value only if it does not resolve
    }

    return $set
}

function Test-IdentityIntersection {
    param(
        [System.Collections.Generic.HashSet[string]]$Left,
        [System.Collections.Generic.HashSet[string]]$Right
    )

    foreach ($item in $Left) {
        if ($Right.Contains($item)) {
            return $true
        }
    }

    return $false
}

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [string]$MailboxDisplayName,
        [string]$MailboxAddress,
        [string]$MailboxType,
        [string]$AccessType,
        [string]$GrantedVia,
        [string]$MatchedTrustee
    )

    $Results.Add([pscustomobject]@{
        MailboxDisplayName = $MailboxDisplayName
        MailboxAddress     = $MailboxAddress
        MailboxType        = $MailboxType
        AccessType         = $AccessType
        GrantedVia         = $GrantedVia
        MatchedTrustee     = $MatchedTrustee
    })
}

Write-Step "Preparing output folder"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Step "Loading modules"
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop

Write-Step "Connecting if needed"
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Directory.Read.All" | Out-Null
}

Write-Step "Resolving target user"
$user = Get-MgUser -UserId $UserEmail -Property Id,DisplayName,UserPrincipalName,Mail
if (-not $user) {
    throw "Could not resolve user: $UserEmail"
}

$targetIdentitySet = New-IdentitySet
Add-NormalizedValue -Set $targetIdentitySet -Value $user.UserPrincipalName
Add-NormalizedValue -Set $targetIdentitySet -Value $user.Mail
Add-NormalizedValue -Set $targetIdentitySet -Value $user.DisplayName
Add-NormalizedValue -Set $targetIdentitySet -Value $UserEmail

Write-Step "Resolving Exchange recipient details for target user"
try {
    $targetRecipient = Get-Recipient -Identity $user.UserPrincipalName -ErrorAction Stop
    $recipientStrings = Expand-RecipientIdentityStrings -Recipient $targetRecipient
    Merge-HashSet -Target $targetIdentitySet -Source $recipientStrings
}
catch {
    Write-Warning "Could not resolve Exchange recipient details for $($user.UserPrincipalName). Continuing with Microsoft Graph identity values only."
}

Write-Step "Collecting group memberships"
$groupIdentitySets = @()

try {
    $memberOf = Get-MgUserMemberOf -UserId $user.Id -All

    foreach ($obj in $memberOf) {
        $odataType = $null

        if ($obj.AdditionalProperties -and $obj.AdditionalProperties.ContainsKey('@odata.type')) {
            $odataType = [string]$obj.AdditionalProperties['@odata.type']
        }

        if ($odataType -eq '#microsoft.graph.group') {
            $groupId = $obj.Id
            if ($groupId) {
                try {
                    $g = Get-MgGroup -GroupId $groupId -Property Id,DisplayName,Mail

                    $groupSet = New-IdentitySet
                    Add-NormalizedValue -Set $groupSet -Value $g.Id
                    Add-NormalizedValue -Set $groupSet -Value $g.DisplayName
                    Add-NormalizedValue -Set $groupSet -Value $g.Mail

                    if (-not [string]::IsNullOrWhiteSpace($g.Mail)) {
                        try {
                            $groupRecipient = Get-Recipient -Identity $g.Mail -ErrorAction Stop
                            $groupRecipientStrings = Expand-RecipientIdentityStrings -Recipient $groupRecipient
                            Merge-HashSet -Target $groupSet -Source $groupRecipientStrings
                        }
                        catch {
                            # Not every group is mail-enabled in Exchange Online
                        }
                    }

                    $groupIdentitySets += $groupSet
                }
                catch {
                    # Skip groups that cannot be resolved
                }
            }
        }
    }
}
catch {
    Write-Warning "Could not enumerate group memberships. Group-based mailbox access may be incomplete."
}

Write-Step "Enumerating mailboxes"
if ($SharedMailboxesOnly) {
    # Shared mailboxes only, faster in large environments
    $allMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails
}
else {
    # All mailboxes
    $allMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails
}

$results = New-Object 'System.Collections.Generic.List[object]'

$total = @($allMailboxes).Count
if ($total -gt 0) {
    Write-Progress -Id 1 -Activity "Scanning mailboxes" -Status ("0 of {0}" -f $total) -PercentComplete 0
}

for ($i = 0; $i -lt $total; $i++) {
    $mbx = $allMailboxes[$i]

    $mailboxAddress = Get-SafeString $mbx.PrimarySmtpAddress
    $mailboxName = Get-SafeString $mbx.DisplayName
    $mailboxType = Get-SafeString $mbx.RecipientTypeDetails

    $pct = [int]((($i + 1) / [double]$total) * 100)
    $statusName = if ([string]::IsNullOrWhiteSpace($mailboxAddress)) { $mbx.Identity } else { $mailboxAddress }
    Write-Progress -Id 1 -Activity "Scanning mailboxes" -Status ("{0} of {1}: {2}" -f ($i + 1), $total, $statusName) -PercentComplete $pct

    # Full Access
    try {
        $faEntries = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction Stop | Where-Object {
            -not $_.IsInherited -and
            -not $_.Deny -and
            ($_.AccessRights -contains 'FullAccess') -and
            ([string]$_.User -notmatch '^NT AUTHORITY\\SELF$')
        }

        foreach ($entry in $faEntries) {
            $trusteeRaw = Get-SafeString $entry.User
            $trusteeSet = Resolve-IdentityStringToSet -IdentityString $trusteeRaw

            if (Test-IdentityIntersection -Left $trusteeSet -Right $targetIdentitySet) {
                Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'FullAccess' -GrantedVia 'Direct' -MatchedTrustee $trusteeRaw
                continue
            }

            foreach ($groupSet in $groupIdentitySets) {
                if (Test-IdentityIntersection -Left $trusteeSet -Right $groupSet) {
                    Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'FullAccess' -GrantedVia 'Group' -MatchedTrustee $trusteeRaw
                    break
                }
            }
        }
    }
    catch {
        Write-Warning "Full Access check failed for $mailboxAddress : $($_.Exception.Message)"
    }

    # Send As
    try {
        $saEntries = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue | Where-Object {
            ($_.AccessRights -contains 'SendAs') -and
            ([string]$_.Trustee -notmatch '^NT AUTHORITY\\SELF$')
        }

        foreach ($entry in $saEntries) {
            $trusteeRaw = Get-SafeString $entry.Trustee
            $trusteeSet = Resolve-IdentityStringToSet -IdentityString $trusteeRaw

            if (Test-IdentityIntersection -Left $trusteeSet -Right $targetIdentitySet) {
                Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'SendAs' -GrantedVia 'Direct' -MatchedTrustee $trusteeRaw
                continue
            }

            foreach ($groupSet in $groupIdentitySets) {
                if (Test-IdentityIntersection -Left $trusteeSet -Right $groupSet) {
                    Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'SendAs' -GrantedVia 'Group' -MatchedTrustee $trusteeRaw
                    break
                }
            }
        }
    }
    catch {
        Write-Warning "Send As check failed for $mailboxAddress : $($_.Exception.Message)"
    }

    # Send on Behalf
    try {
        $fullMailbox = Get-Mailbox -Identity $mbx.Identity -ErrorAction Stop
        $sobDelegates = @($fullMailbox.GrantSendOnBehalfTo)

        foreach ($delegate in $sobDelegates) {
            $delegateRaw = Get-SafeString $delegate
            if ([string]::IsNullOrWhiteSpace($delegateRaw)) { continue }

            $delegateSet = Resolve-IdentityStringToSet -IdentityString $delegateRaw

            if (Test-IdentityIntersection -Left $delegateSet -Right $targetIdentitySet) {
                Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'SendOnBehalf' -GrantedVia 'Direct' -MatchedTrustee $delegateRaw
                continue
            }

            foreach ($groupSet in $groupIdentitySets) {
                if (Test-IdentityIntersection -Left $delegateSet -Right $groupSet) {
                    Add-ResultRow -Results $results -MailboxDisplayName $mailboxName -MailboxAddress $mailboxAddress -MailboxType $mailboxType -AccessType 'SendOnBehalf' -GrantedVia 'Group' -MatchedTrustee $delegateRaw
                    break
                }
            }
        }
    }
    catch {
        Write-Warning "Send on Behalf check failed for $mailboxAddress : $($_.Exception.Message)"
    }
}

try { Write-Progress -Id 1 -Activity "Scanning mailboxes" -Completed } catch {}

Write-Step "De-duplicating and exporting results"
$final = $results | Sort-Object MailboxAddress, AccessType, GrantedVia, MatchedTrustee -Unique

$safeUserPart = ($user.UserPrincipalName -replace '[^a-zA-Z0-9@._-]', '_')
$csvPath = Join-Path $OutputFolder "MailboxAccess_$safeUserPart.csv"
$jsonPath = Join-Path $OutputFolder "MailboxAccess_$safeUserPart.json"

$final | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
$final | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8

Write-Step "Results"
$final | Format-Table -AutoSize

Write-Host ""
Write-Host "CSV  : $csvPath" -ForegroundColor Green
Write-Host "JSON : $jsonPath" -ForegroundColor Green
Write-Host "Found $($final.Count) mailbox access entries for $($user.UserPrincipalName)." -ForegroundColor Yellow