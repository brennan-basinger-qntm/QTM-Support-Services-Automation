[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$UserPrincipalName,
    [string]$LogPath = ".\OffboardingAccessAudit.csv",
    [switch]$UseDeviceAuthentication,
    [switch]$AllowRoleAssignableGroups,
    [switch]$EnableRemoval
)

$ErrorActionPreference = 'Stop'
$script:ToolCmdlet = $PSCmdlet
$script:UseDeviceAuthentication = [bool]$UseDeviceAuthentication
$script:EnableRemoval = [bool]$EnableRemoval
$script:AllowRoleAssignableGroups = [bool]$AllowRoleAssignableGroups
$script:GraphConnected = $false
$script:ExchangeConnected = $false
$script:ExchangeConnectedAs = $null
$script:MailboxRecipientTypeDetails = @('SharedMailbox','UserMailbox','RoomMailbox','EquipmentMailbox')

$script:GraphReadScopes = @(
    'User.Read.All',
    'Group.Read.All'
)

$script:GraphWriteScopes = @(
    'User.Read.All',
    'Group.Read.All',
    'GroupMember.ReadWrite.All'
)

if ($AllowRoleAssignableGroups) {
    $script:GraphWriteScopes += 'RoleManagement.ReadWrite.Directory'
}

$requiredGraphCmdlets = @(
    'Connect-MgGraph',
    'Get-MgContext',
    'Get-MgUser',
    'Get-MgUserMemberOf',
    'Get-MgUserTransitiveMemberOf',
    'Get-MgGroup',
    'Remove-MgGroupMemberByRef'
)

$requiredExchangeCmdlets = @(
    'Connect-ExchangeOnline',
    'Get-Recipient',
    'Get-DistributionGroup',
    'Get-DistributionGroupMember',
    'Get-EXOMailbox',
    'Get-EXOMailboxPermission',
    'Get-EXORecipientPermission',
    'Remove-DistributionGroupMember',
    'Set-DistributionGroup',
    'Remove-MailboxPermission',
    'Remove-RecipientPermission',
    'Set-Mailbox'
)

foreach ($cmd in $requiredGraphCmdlets) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "Required cmdlet '$cmd' was not found. Install Microsoft Graph PowerShell first: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
}

foreach ($cmd in $requiredExchangeCmdlets) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "Required cmdlet '$cmd' was not found. Install Exchange Online PowerShell first: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
}

function Normalize-Value {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().ToLowerInvariant()
}

function Resolve-KeySet {
    param(
        [Parameter(Mandatory)]
        $Set
    )

    $resolved = $null

    if ($Set -is [System.Collections.Generic.HashSet[string]]) {
        $resolved = $Set
    }
    elseif ($Set -is [System.Array] -and $Set.Count -eq 1) {
        $resolved = Resolve-KeySet -Set $Set[0]
    }
    elseif ($Set -and $Set.PSObject.Properties.Name -contains 'KeySet' -and $Set.KeySet -is [System.Collections.Generic.HashSet[string]]) {
        $resolved = $Set.KeySet
    }
    elseif ($null -eq $Set) {
        throw 'Internal error: key set is null.'
    }
    else {
        $typeName = if ($Set -and $Set.GetType()) { $Set.GetType().FullName } else { 'unknown' }
        throw "Internal error: expected HashSet[string] but received '$typeName'."
    }

    Write-Output -NoEnumerate $resolved
}

function Add-NormalizedValue {
    param(
        [Parameter(Mandatory)]
        $Set,
        $Value
    )

    $resolvedSet = Resolve-KeySet -Set $Set
    $normalized = Normalize-Value -Value $Value
    if ($normalized) {
        [void]$resolvedSet.Add($normalized)
    }
}

function New-KeySet {
    [pscustomobject]@{
        KeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
}

function Ensure-LogPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $parent = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrWhiteSpace($parent)) {
        return
    }

    if (-not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Write-StageMessage {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host ''
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor Cyan
}

function Write-ToolProgress {
    param(
        [Parameter(Mandatory)]
        [string]$Activity,
        [string]$Status = '',
        [double]$PercentComplete = 0,
        [int]$Id = 1,
        [switch]$Completed
    )

    if ($Completed) {
        Write-Progress -Id $Id -Activity $Activity -Completed
        return
    }

    Write-Progress -Id $Id -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

function Write-AuditRows {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return
    }

    Ensure-LogPath -Path $LogPath

    if (Test-Path -Path $LogPath) {
        $Rows | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $Rows | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    }
}

function Get-GraphExecutionIdentity {
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            return $ctx.Account
        }
    }
    catch {}

    return $null
}

function Get-ExchangeExecutionIdentity {
    if ($script:ExchangeConnectedAs) {
        return $script:ExchangeConnectedAs
    }

    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $conn = Get-ConnectionInformation | Select-Object -First 1
            if ($conn -and $conn.UserPrincipalName) {
                $script:ExchangeConnectedAs = $conn.UserPrincipalName
                return $script:ExchangeConnectedAs
            }
        }
        catch {}
    }

    return $null
}

function Ensure-GraphConnection {
    param(
        [Parameter(Mandatory)]
        [string[]]$Scopes
    )

    $ctx = $null
    try {
        $ctx = Get-MgContext -ErrorAction Stop
    }
    catch {
        $ctx = $null
    }

    $needsConnect = $true
    if ($ctx -and $ctx.Scopes) {
        $missingScopes = $Scopes | Where-Object { $ctx.Scopes -notcontains $_ }
        if (-not $missingScopes -or $missingScopes.Count -eq 0) {
            $needsConnect = $false
        }
    }

    if ($needsConnect) {
        if ($script:UseDeviceAuthentication) {
            Connect-MgGraph -Scopes $Scopes -NoWelcome -UseDeviceAuthentication | Out-Null
        }
        else {
            Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
        }
    }

    $script:GraphConnected = $true
}

function Ensure-ExchangeConnection {
    if ($script:ExchangeConnected) {
        return
    }

    if ($script:UseDeviceAuthentication) {
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            throw 'Exchange Online device authentication requires PowerShell 7 or later. Either run this script in PowerShell 7 or omit -UseDeviceAuthentication.'
        }

        Connect-ExchangeOnline -Device -ShowBanner:$false | Out-Null
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    }

    $script:ExchangeConnected = $true
    $null = Get-ExchangeExecutionIdentity
}

function Get-ODataType {
    param($DirectoryObject)

    if ($null -eq $DirectoryObject) {
        return $null
    }

    if ($DirectoryObject.PSObject.Properties.Name -contains '@odata.type') {
        return $DirectoryObject.'@odata.type'
    }

    if ($DirectoryObject.PSObject.Properties.Name -contains 'OdataType' -and $DirectoryObject.OdataType) {
        return $DirectoryObject.OdataType
    }

    if ($DirectoryObject.AdditionalProperties -and $DirectoryObject.AdditionalProperties.ContainsKey('@odata.type')) {
        return $DirectoryObject.AdditionalProperties['@odata.type']
    }

    return $null
}

function Get-EntraGroupCategory {
    param(
        [Parameter(Mandatory)]
        $Group
    )

    $groupTypes = @($Group.GroupTypes)

    if ($groupTypes -contains 'Unified') {
        return 'Microsoft 365'
    }

    if ($Group.SecurityEnabled -and -not $Group.MailEnabled) {
        return 'Security'
    }

    if ($Group.SecurityEnabled -and $Group.MailEnabled) {
        return 'Mail-enabled Security'
    }

    if (-not $Group.SecurityEnabled -and $Group.MailEnabled) {
        return 'Distribution'
    }

    return 'Other'
}

function Test-EntraGroupRemovalEligibility {
    param(
        [Parameter(Mandatory)]
        $InventoryRow
    )

    if ($InventoryRow.Path -ne 'Direct') {
        return [pscustomobject]@{
            CanRemove = $false
            Reason    = 'Transitive membership. Remove the user from the parent group instead.'
        }
    }

    if ($InventoryRow.ObjectType -notin @('Security', 'Microsoft 365')) {
        return [pscustomobject]@{
            CanRemove = $false
            Reason    = "Group type '$($InventoryRow.ObjectType)' is read-only through Microsoft Graph."
        }
    }

    if ($InventoryRow.DynamicMembership) {
        return [pscustomobject]@{
            CanRemove = $false
            Reason    = 'Dynamic membership group.'
        }
    }

    if ($InventoryRow.OnPremisesSyncEnabled) {
        return [pscustomobject]@{
            CanRemove = $false
            Reason    = 'Group is synchronized from on-premises.'
        }
    }

    if ($InventoryRow.RoleAssignable -and -not $script:AllowRoleAssignableGroups) {
        return [pscustomobject]@{
            CanRemove = $false
            Reason    = 'Role-assignable group blocked. Re-run with -AllowRoleAssignableGroups if appropriate.'
        }
    }

    return [pscustomobject]@{
        CanRemove = $true
        Reason    = ''
    }
}

function Get-DirectoryUser {
    param(
        [Parameter(Mandatory)]
        [string]$UPN
    )

    Ensure-GraphConnection -Scopes $script:GraphReadScopes
    return Get-MgUser -UserId $UPN -Property Id,DisplayName,UserPrincipalName,AccountEnabled
}

function Get-GroupById {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    return Get-MgGroup -GroupId $GroupId -Property Id,DisplayName,Description,Mail,MailEnabled,SecurityEnabled,GroupTypes,MembershipRule,MembershipRuleProcessingState,OnPremisesSyncEnabled,IsAssignableToRole
}

function Get-ExchangeRecipientRecord {
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    Ensure-ExchangeConnection
    return Get-Recipient -Identity $Identity -ErrorAction Stop
}

function Resolve-TargetContext {
    param(
        [Parameter(Mandatory)]
        [string]$UPN
    )

    $context = [ordered]@{
        InputUpn          = $UPN
        UserPrincipalName = $UPN
        EntraUser         = $null
        ExchangeRecipient = $null
    }

    Write-StageMessage -Message ("Resolving target user '{0}' in Entra ID and Exchange Online..." -f $UPN)

    try {
        Write-ToolProgress -Id 2 -Activity 'Resolving target user' -Status 'Looking up user in Entra ID' -PercentComplete 35
        $context.EntraUser = Get-DirectoryUser -UPN $UPN
        $context.UserPrincipalName = $context.EntraUser.UserPrincipalName
    }
    catch {
        Write-Warning "Entra ID lookup failed for '$UPN'. $($_.Exception.Message)"
    }

    try {
        Write-ToolProgress -Id 2 -Activity 'Resolving target user' -Status 'Looking up user in Exchange Online' -PercentComplete 75
        $context.ExchangeRecipient = Get-ExchangeRecipientRecord -Identity $UPN
        if (-not $context.EntraUser -and $context.ExchangeRecipient.PrimarySmtpAddress) {
            $context.UserPrincipalName = [string]$context.ExchangeRecipient.PrimarySmtpAddress
        }
    }
    catch {
        Write-Warning "Exchange Online lookup failed for '$UPN'. $($_.Exception.Message)"
    }
    finally {
        Write-ToolProgress -Id 2 -Activity 'Resolving target user' -Completed
    }

    if (-not $context.EntraUser -and -not $context.ExchangeRecipient) {
        throw "Target '$UPN' was not found in Entra ID or Exchange Online."
    }

    return [pscustomobject]$context
}

function Get-TargetKeySet {
    param(
        [Parameter(Mandatory)]
        $TargetContext
    )

    $keys = New-KeySet

    Add-NormalizedValue -Set $keys -Value $TargetContext.InputUpn
    Add-NormalizedValue -Set $keys -Value $TargetContext.UserPrincipalName

    if ($TargetContext.EntraUser) {
        Add-NormalizedValue -Set $keys -Value $TargetContext.EntraUser.UserPrincipalName
        Add-NormalizedValue -Set $keys -Value $TargetContext.EntraUser.Id
    }

    if ($TargetContext.ExchangeRecipient) {
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.PrimarySmtpAddress
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.WindowsEmailAddress
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.ExternalDirectoryObjectId
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.Guid
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.DistinguishedName
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.Alias
        Add-NormalizedValue -Set $keys -Value $TargetContext.ExchangeRecipient.UserPrincipalName
    }

    return $keys
}

function Test-ObjectMatchesTarget {
    param(
        [Parameter(Mandatory)]
        $Object,
        [Parameter(Mandatory)]
        $TargetKeys
    )

    $TargetKeys = Resolve-KeySet -Set $TargetKeys

    $candidateProperties = @(
        'PrimarySmtpAddress',
        'WindowsEmailAddress',
        'WindowsLiveID',
        'UserPrincipalName',
        'ExternalDirectoryObjectId',
        'Guid',
        'DistinguishedName',
        'Alias'
    )

    foreach ($property in $candidateProperties) {
        if ($Object.PSObject.Properties.Name -contains $property) {
            $normalized = Normalize-Value -Value $Object.$property
            if ($normalized -and $TargetKeys.Contains($normalized)) {
                return $true
            }
        }
    }

    $stringValue = Normalize-Value -Value ($Object.ToString())
    if ($stringValue) {
        $looksLikeUniqueIdentifier = ($stringValue -match '@') -or ($stringValue -match '=') -or ($stringValue -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')
        if ($looksLikeUniqueIdentifier -and $TargetKeys.Contains($stringValue)) {
            return $true
        }
    }

    return $false
}

function New-AuditRow {
    param(
        [Parameter(Mandatory)]
        [string]$Operation,
        [Parameter(Mandatory)]
        [string]$Provider,
        [Parameter(Mandatory)]
        [string]$Category,
        [Parameter(Mandatory)]
        $TargetContext,
        [Parameter(Mandatory)]
        [string]$ObjectDisplayName,
        [string]$ObjectIdentity,
        [string]$ObjectAddress,
        [string]$ObjectType,
        [string]$Path,
        [string]$Permission,
        [bool]$DynamicMembership = $false,
        [bool]$OnPremisesSyncEnabled = $false,
        [bool]$RoleAssignable = $false,
        [bool]$Removable = $false,
        [bool]$Transferable = $false,
        [int]$OwnerCount = 0,
        [string]$Status = 'Listed',
        [string]$Reason = ''
    )

    $executedBy = if ($Provider -eq 'Entra') { Get-GraphExecutionIdentity } else { Get-ExchangeExecutionIdentity }

    [pscustomobject]@{
        Timestamp             = (Get-Date).ToString('s')
        Operation             = $Operation
        Provider              = $Provider
        Category              = $Category
        UserDisplayName       = if ($TargetContext.EntraUser) { $TargetContext.EntraUser.DisplayName } elseif ($TargetContext.ExchangeRecipient) { $TargetContext.ExchangeRecipient.DisplayName } else { $TargetContext.UserPrincipalName }
        UserPrincipalName     = $TargetContext.UserPrincipalName
        ObjectDisplayName     = $ObjectDisplayName
        ObjectIdentity        = $ObjectIdentity
        ObjectAddress         = $ObjectAddress
        ObjectType            = $ObjectType
        Path                  = $Path
        Permission            = $Permission
        DynamicMembership     = [bool]$DynamicMembership
        OnPremisesSyncEnabled = [bool]$OnPremisesSyncEnabled
        RoleAssignable        = [bool]$RoleAssignable
        Removable             = [bool]$Removable
        Transferable          = [bool]$Transferable
        OwnerCount            = [int]$OwnerCount
        Status                = $Status
        Reason                = $Reason
        ExecutedBy            = $executedBy
    }
}

function Get-EntraGroupInventory {
    param(
        [Parameter(Mandatory)]
        $TargetContext,
        [switch]$IncludeTransitive
    )

    if (-not $TargetContext.EntraUser) {
        throw 'The target could not be resolved in Entra ID.'
    }

    Write-StageMessage -Message ("Enumerating Entra {0}group memberships for {1}..." -f $(if ($IncludeTransitive) { 'effective ' } else { 'direct ' }), $TargetContext.UserPrincipalName)
    Ensure-GraphConnection -Scopes $script:GraphReadScopes

    $directRefs = @(Get-MgUserMemberOf -UserId $TargetContext.EntraUser.Id -All)
    $directGroupIds = Resolve-KeySet -Set (New-KeySet)
    foreach ($ref in $directRefs) {
        if ((Get-ODataType -DirectoryObject $ref) -eq '#microsoft.graph.group') {
            Add-NormalizedValue -Set $directGroupIds -Value $ref.Id
        }
    }

    $membershipRefs = $directRefs
    if ($IncludeTransitive) {
        $membershipRefs = @(Get-MgUserTransitiveMemberOf -UserId $TargetContext.EntraUser.Id -All)
    }

    $uniqueGroupIds = [ordered]@{}
    foreach ($ref in $membershipRefs) {
        if ((Get-ODataType -DirectoryObject $ref) -eq '#microsoft.graph.group') {
            $uniqueGroupIds[$ref.Id] = $true
        }
    }

    $groupIds = @($uniqueGroupIds.Keys)
    $total = [Math]::Max($groupIds.Count, 1)
    $index = 0

    $rows = foreach ($groupId in $groupIds) {
        $index++
        Write-ToolProgress -Id 11 -Activity 'Scanning Entra groups' -Status ("{0}/{1}: {2}" -f $index, $groupIds.Count, $groupId) -PercentComplete (($index / $total) * 100)

        $group = Get-GroupById -GroupId $groupId
        $groupTypes = @($group.GroupTypes)
        $isDynamic = ($groupTypes -contains 'DynamicMembership') -or (-not [string]::IsNullOrWhiteSpace($group.MembershipRule))
        $path = if ($directGroupIds.Contains((Normalize-Value -Value $group.Id))) { 'Direct' } else { 'Transitive' }

        $row = New-AuditRow -Operation 'Inventory' -Provider 'Entra' -Category 'GroupMembership' -TargetContext $TargetContext -ObjectDisplayName $group.DisplayName -ObjectIdentity $group.Id -ObjectAddress ([string]$group.Mail) -ObjectType (Get-EntraGroupCategory -Group $group) -Path $path -Permission 'Member' -DynamicMembership:$isDynamic -OnPremisesSyncEnabled:([bool]$group.OnPremisesSyncEnabled) -RoleAssignable:([bool]$group.IsAssignableToRole)

        $eligibility = Test-EntraGroupRemovalEligibility -InventoryRow $row
        $row.Removable = [bool]$eligibility.CanRemove
        $row.Reason = $eligibility.Reason
        $row
    }

    Write-ToolProgress -Id 11 -Activity 'Scanning Entra groups' -Completed
    return @($rows | Sort-Object Path, ObjectDisplayName)
}

function Test-ExchangeGroupActionEligibility {
    param(
        [Parameter(Mandatory)]
        $Group,
        [ValidateSet('Membership','Ownership')]
        [string]$Action
    )

    $isDirSynced = $false
    if ($Group.PSObject.Properties.Name -contains 'IsDirSynced' -and $null -ne $Group.IsDirSynced) {
        $isDirSynced = [bool]$Group.IsDirSynced
    }

    if ($isDirSynced) {
        return [pscustomobject]@{
            Allowed = $false
            Reason  = 'Exchange group is directory-synchronized. Manage it from on-premises.'
        }
    }

    if ($Action -eq 'Ownership' -and @($Group.ManagedBy).Count -lt 1) {
        return [pscustomobject]@{
            Allowed = $false
            Reason  = 'No current group owner data was returned.'
        }
    }

    return [pscustomobject]@{
        Allowed = $true
        Reason  = ''
    }
}

function Get-DistributionListInventory {
    param(
        [Parameter(Mandatory)]
        $TargetContext
    )

    if (-not $TargetContext.ExchangeRecipient) {
        throw 'The target could not be resolved in Exchange Online.'
    }

    Write-StageMessage -Message ("Enumerating distribution list membership and ownership for {0}..." -f $TargetContext.UserPrincipalName)
    Ensure-ExchangeConnection

    $targetKeys = Resolve-KeySet -Set (Get-TargetKeySet -TargetContext $TargetContext)
    $groups = @(Get-DistributionGroup -ResultSize Unlimited)
    $rows = New-Object System.Collections.Generic.List[object]

    $total = [Math]::Max($groups.Count, 1)
    $index = 0

    foreach ($group in $groups) {
        $index++
        Write-ToolProgress -Id 12 -Activity 'Scanning distribution lists' -Status ("{0}/{1}: {2}" -f $index, $groups.Count, $group.DisplayName) -PercentComplete (($index / $total) * 100)

        $ownerMatch = $false
        foreach ($owner in @($group.ManagedBy)) {
            if (Test-ObjectMatchesTarget -Object $owner -TargetKeys $targetKeys) {
                $ownerMatch = $true
                break
            }
        }

        if ($ownerMatch) {
            $eligibility = Test-ExchangeGroupActionEligibility -Group $group -Action Ownership
            $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'DistributionOwnership' -TargetContext $TargetContext -ObjectDisplayName $group.DisplayName -ObjectIdentity ([string]$group.Identity) -ObjectAddress ([string]$group.PrimarySmtpAddress) -ObjectType ([string]$group.RecipientTypeDetails) -Path 'Direct' -Permission 'Owner' -OnPremisesSyncEnabled:([bool]($group.PSObject.Properties.Name -contains 'IsDirSynced' -and $group.IsDirSynced)) -Removable:$false -Transferable:([bool]$eligibility.Allowed) -OwnerCount (@($group.ManagedBy).Count) -Reason $eligibility.Reason
            [void]$rows.Add($row)
        }

        $memberMatch = $false
        try {
            $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
            foreach ($member in $members) {
                if (Test-ObjectMatchesTarget -Object $member -TargetKeys $targetKeys) {
                    $memberMatch = $true
                    break
                }
            }
        }
        catch {
            $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'DistributionMembership' -TargetContext $TargetContext -ObjectDisplayName $group.DisplayName -ObjectIdentity ([string]$group.Identity) -ObjectAddress ([string]$group.PrimarySmtpAddress) -ObjectType ([string]$group.RecipientTypeDetails) -Path 'Direct' -Permission 'Member' -Removable:$false -Reason ("Unable to read membership. {0}" -f $_.Exception.Message)
            [void]$rows.Add($row)
            continue
        }

        if ($memberMatch) {
            $eligibility = Test-ExchangeGroupActionEligibility -Group $group -Action Membership
            $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'DistributionMembership' -TargetContext $TargetContext -ObjectDisplayName $group.DisplayName -ObjectIdentity ([string]$group.Identity) -ObjectAddress ([string]$group.PrimarySmtpAddress) -ObjectType ([string]$group.RecipientTypeDetails) -Path 'Direct' -Permission 'Member' -OnPremisesSyncEnabled:([bool]($group.PSObject.Properties.Name -contains 'IsDirSynced' -and $group.IsDirSynced)) -Removable:([bool]$eligibility.Allowed) -Reason $eligibility.Reason
            [void]$rows.Add($row)
        }
    }

    Write-ToolProgress -Id 12 -Activity 'Scanning distribution lists' -Completed
    return @($rows | Sort-Object Category, ObjectDisplayName)
}

function Get-MailboxPermissionInventory {
    param(
        [Parameter(Mandatory)]
        $TargetContext
    )

    if (-not $TargetContext.ExchangeRecipient) {
        throw 'The target could not be resolved in Exchange Online.'
    }

    Write-StageMessage -Message ("Enumerating mailbox delegate access for {0}..." -f $TargetContext.UserPrincipalName)
    Ensure-ExchangeConnection

    $rows = New-Object System.Collections.Generic.List[object]
    $targetKeys = Resolve-KeySet -Set (Get-TargetKeySet -TargetContext $TargetContext)
    $delegateIdentifier = if ($TargetContext.ExchangeRecipient.PrimarySmtpAddress) { [string]$TargetContext.ExchangeRecipient.PrimarySmtpAddress } else { $TargetContext.UserPrincipalName }
    $delegateForMailboxPermission = if ($TargetContext.ExchangeRecipient.DistinguishedName) { [string]$TargetContext.ExchangeRecipient.DistinguishedName } else { $delegateIdentifier }

    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo)
    $total = [Math]::Max($mailboxes.Count, 1)
    $index = 0

    foreach ($mailbox in $mailboxes) {
        $index++
        Write-ToolProgress -Id 13 -Activity 'Scanning mailbox permissions' -Status ("{0}/{1}: {2}" -f $index, $mailboxes.Count, $mailbox.DisplayName) -PercentComplete (($index / $total) * 100)

        if ($mailbox.RecipientTypeDetails -notin $script:MailboxRecipientTypeDetails) {
            continue
        }

        try {
            $fullAccessEntries = @(Get-EXOMailboxPermission -Identity $mailbox.UserPrincipalName -User $delegateForMailboxPermission -ErrorAction Stop)
            foreach ($entry in $fullAccessEntries) {
                if ($entry.Deny -or $entry.IsInherited) {
                    continue
                }

                if (@($entry.AccessRights) -contains 'FullAccess') {
                    $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'MailboxPermission' -TargetContext $TargetContext -ObjectDisplayName $mailbox.DisplayName -ObjectIdentity ([string]$mailbox.UserPrincipalName) -ObjectAddress ([string]$mailbox.PrimarySmtpAddress) -ObjectType ([string]$mailbox.RecipientTypeDetails) -Path 'Direct' -Permission 'FullAccess' -Removable:$true
                    [void]$rows.Add($row)
                }
            }
        }
        catch {}

        $sendOnBehalfEntries = @($mailbox.GrantSendOnBehalfTo)
        $sendOnBehalfMatch = $false
        foreach ($delegate in $sendOnBehalfEntries) {
            if (Test-ObjectMatchesTarget -Object $delegate -TargetKeys $targetKeys) {
                $sendOnBehalfMatch = $true
                break
            }
        }

        if ($sendOnBehalfMatch) {
            $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'MailboxPermission' -TargetContext $TargetContext -ObjectDisplayName $mailbox.DisplayName -ObjectIdentity ([string]$mailbox.UserPrincipalName) -ObjectAddress ([string]$mailbox.PrimarySmtpAddress) -ObjectType ([string]$mailbox.RecipientTypeDetails) -Path 'Direct' -Permission 'SendOnBehalf' -Removable:$true
            [void]$rows.Add($row)
        }
    }

    Write-ToolProgress -Id 13 -Activity 'Scanning mailbox permissions' -Completed

    Write-StageMessage -Message ("Querying SendAs delegate permissions for {0}..." -f $TargetContext.UserPrincipalName)

    try {
        $sendAsEntries = @(Get-EXORecipientPermission -Trustee $delegateIdentifier -ResultSize Unlimited -ErrorAction Stop)
        foreach ($entry in $sendAsEntries) {
            $recipient = $null
            try {
                $recipient = Get-Recipient -Identity $entry.Identity -ErrorAction Stop
            }
            catch {
                continue
            }

            if ($recipient.RecipientTypeDetails -notin $script:MailboxRecipientTypeDetails) {
                continue
            }

            $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'MailboxPermission' -TargetContext $TargetContext -ObjectDisplayName $recipient.DisplayName -ObjectIdentity ([string]$recipient.Identity) -ObjectAddress ([string]$recipient.PrimarySmtpAddress) -ObjectType ([string]$recipient.RecipientTypeDetails) -Path 'Direct' -Permission 'SendAs' -Removable:$true
            [void]$rows.Add($row)
        }
    }
    catch {
        $row = New-AuditRow -Operation 'Inventory' -Provider 'Exchange' -Category 'MailboxPermission' -TargetContext $TargetContext -ObjectDisplayName '<query>' -ObjectIdentity '<query>' -ObjectAddress '' -ObjectType 'Mailbox' -Path 'Direct' -Permission 'SendAs' -Removable:$false -Reason ("Unable to query SendAs permissions. {0}" -f $_.Exception.Message)
        [void]$rows.Add($row)
    }

    return @($rows | Sort-Object ObjectDisplayName, Permission -Unique)
}

function Get-CombinedOffboardingReport {
    param(
        [Parameter(Mandatory)]
        $TargetContext
    )

    Write-StageMessage -Message ("Building combined offboarding report for {0}..." -f $TargetContext.UserPrincipalName)
    $allRows = New-Object System.Collections.Generic.List[object]

    if ($TargetContext.EntraUser) {
        foreach ($row in (Get-EntraGroupInventory -TargetContext $TargetContext)) {
            [void]$allRows.Add($row)
        }
    }

    if ($TargetContext.ExchangeRecipient) {
        foreach ($row in (Get-DistributionListInventory -TargetContext $TargetContext)) {
            [void]$allRows.Add($row)
        }

        foreach ($row in (Get-MailboxPermissionInventory -TargetContext $TargetContext)) {
            [void]$allRows.Add($row)
        }
    }

    return @($allRows)
}

function Show-EntraInventory {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [string]$Title = 'Entra group inventory'
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('-' * $Title.Length) -ForegroundColor Cyan
    $Rows | Select-Object ObjectDisplayName, ObjectType, Path, DynamicMembership, OnPremisesSyncEnabled, RoleAssignable, Removable, Reason | Format-Table -AutoSize
    Write-Host ''
    Write-Host ("Total rows: {0}" -f $Rows.Count)
    Write-Host ("Log path : {0}" -f $LogPath)
}

function Show-DistributionInventory {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [string]$Title = 'Distribution list inventory'
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('-' * $Title.Length) -ForegroundColor Cyan
    $Rows | Select-Object ObjectDisplayName, Permission, ObjectAddress, ObjectType, OwnerCount, OnPremisesSyncEnabled, Removable, Transferable, Reason | Format-Table -AutoSize
    Write-Host ''
    Write-Host ("Total rows: {0}" -f $Rows.Count)
    Write-Host ("Log path : {0}" -f $LogPath)
}

function Show-MailboxInventory {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [string]$Title = 'Mailbox permission inventory'
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('-' * $Title.Length) -ForegroundColor Cyan
    $Rows | Select-Object ObjectDisplayName, Permission, ObjectAddress, ObjectType, Removable, Reason | Format-Table -AutoSize
    Write-Host ''
    Write-Host ("Total rows: {0}" -f $Rows.Count)
    Write-Host ("Log path : {0}" -f $LogPath)
}

function Show-CombinedInventory {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [string]$Title = 'Combined offboarding inventory'
    )

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('-' * $Title.Length) -ForegroundColor Cyan
    $Rows | Select-Object Provider, Category, ObjectDisplayName, Permission, ObjectAddress, ObjectType, Path, Removable, Transferable, Reason | Format-Table -AutoSize
    Write-Host ''
    Write-Host ("Total rows: {0}" -f $Rows.Count)
    Write-Host ("Log path : {0}" -f $LogPath)
}

function Select-InventoryRows {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [switch]$EligibleOnly,
        [switch]$TransferOnly
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Warning 'No rows are available to select.'
        return @()
    }

    $search = Read-Host 'Enter part of the object name to search (leave blank for all)'
    $filtered = $Rows | Where-Object { $_.ObjectDisplayName -like "*$search*" }
    if ($EligibleOnly) {
        $filtered = $filtered | Where-Object { $_.Removable }
    }
    if ($TransferOnly) {
        $filtered = $filtered | Where-Object { $_.Transferable }
    }

    if (-not $filtered -or $filtered.Count -eq 0) {
        Write-Warning 'No matching rows were found.'
        return @()
    }

    $index = 0
    $pickList = foreach ($row in $filtered) {
        $index++
        [pscustomobject]@{
            Index             = $index
            SelectionKey      = ('{0}|{1}|{2}' -f $row.ObjectIdentity, $row.Category, $row.Permission)
            ObjectDisplayName = $row.ObjectDisplayName
            Permission        = $row.Permission
            ObjectAddress     = $row.ObjectAddress
            ObjectType        = $row.ObjectType
            Removable         = $row.Removable
            Transferable      = $row.Transferable
            Reason            = $row.Reason
            ObjectIdentity    = $row.ObjectIdentity
            Category          = $row.Category
        }
    }

    Write-Host ''
    Write-Host 'Selectable rows' -ForegroundColor Yellow
    Write-Host '--------------' -ForegroundColor Yellow
    $pickList | Format-Table -AutoSize

    $selection = Read-Host 'Enter one or more index values separated by commas, or A for all eligible matches'
    if ([string]::IsNullOrWhiteSpace($selection)) {
        return @()
    }

    if ($selection.Trim().ToUpperInvariant() -eq 'A') {
        $eligible = if ($TransferOnly) { $pickList | Where-Object { $_.Transferable } } elseif ($EligibleOnly) { $pickList | Where-Object { $_.Removable } } else { $pickList }
        $keys = @($eligible | Select-Object -ExpandProperty SelectionKey)
        return @($Rows | Where-Object { $keys -contains ('{0}|{1}|{2}' -f $_.ObjectIdentity, $_.Category, $_.Permission) })
    }

    $indexes = @()
    foreach ($item in ($selection -split ',')) {
        $number = 0
        if ([int]::TryParse($item.Trim(), [ref]$number)) {
            $indexes += $number
        }
    }

    if (-not $indexes -or $indexes.Count -eq 0) {
        Write-Warning 'No valid indexes were provided.'
        return @()
    }

    $keys = @($pickList | Where-Object { $indexes -contains $_.Index } | Select-Object -ExpandProperty SelectionKey)
    return @($Rows | Where-Object { $keys -contains ('{0}|{1}|{2}' -f $_.ObjectIdentity, $_.Category, $_.Permission) })
}

function Remove-EntraMemberships {
    param(
        [Parameter(Mandatory)]
        $TargetContext,
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if (-not $TargetContext.EntraUser) {
        throw 'The target could not be resolved in Entra ID.'
    }

    Ensure-GraphConnection -Scopes $script:GraphWriteScopes

    $results = foreach ($row in $Rows) {
        $status = 'Skipped'
        $reason = $row.Reason

        if (-not $row.Removable) {
            if ([string]::IsNullOrWhiteSpace($reason)) {
                $reason = 'This membership is not removable by this tool.'
            }
        }
        else {
            try {
                $target = "{0} [{1}]" -f $row.ObjectDisplayName, $row.ObjectIdentity
                $action = "Remove {0} from this Entra group" -f $TargetContext.UserPrincipalName

                if ($script:ToolCmdlet.ShouldProcess($target, $action)) {
                    Remove-MgGroupMemberByRef -GroupId $row.ObjectIdentity -DirectoryObjectId $TargetContext.EntraUser.Id -Confirm:$false -ErrorAction Stop
                    $status = 'Removed'
                    $reason = ''
                }
                else {
                    $reason = 'WhatIf or confirmation prevented removal.'
                }
            }
            catch {
                $status = 'Error'
                $reason = $_.Exception.Message
            }
        }

        New-AuditRow -Operation 'Remove' -Provider 'Entra' -Category $row.Category -TargetContext $TargetContext -ObjectDisplayName $row.ObjectDisplayName -ObjectIdentity $row.ObjectIdentity -ObjectAddress $row.ObjectAddress -ObjectType $row.ObjectType -Path $row.Path -Permission $row.Permission -DynamicMembership:$row.DynamicMembership -OnPremisesSyncEnabled:$row.OnPremisesSyncEnabled -RoleAssignable:$row.RoleAssignable -Removable:$row.Removable -Status $status -Reason $reason
    }

    return @($results)
}

function Remove-DistributionMemberships {
    param(
        [Parameter(Mandatory)]
        $TargetContext,
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if (-not $TargetContext.ExchangeRecipient) {
        throw 'The target could not be resolved in Exchange Online.'
    }

    Ensure-ExchangeConnection
    $memberIdentifier = if ($TargetContext.ExchangeRecipient.PrimarySmtpAddress) { [string]$TargetContext.ExchangeRecipient.PrimarySmtpAddress } else { $TargetContext.UserPrincipalName }

    $results = foreach ($row in $Rows) {
        $status = 'Skipped'
        $reason = $row.Reason

        if (-not $row.Removable) {
            if ([string]::IsNullOrWhiteSpace($reason)) {
                $reason = 'This distribution list membership is not removable by this tool.'
            }
        }
        else {
            try {
                $target = "{0} [{1}]" -f $row.ObjectDisplayName, $row.ObjectIdentity
                $action = "Remove {0} from this distribution list" -f $memberIdentifier

                if ($script:ToolCmdlet.ShouldProcess($target, $action)) {
                    Remove-DistributionGroupMember -Identity $row.ObjectIdentity -Member $memberIdentifier -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                    $status = 'Removed'
                    $reason = ''
                }
                else {
                    $reason = 'WhatIf or confirmation prevented removal.'
                }
            }
            catch {
                $status = 'Error'
                $reason = $_.Exception.Message
            }
        }

        New-AuditRow -Operation 'Remove' -Provider 'Exchange' -Category $row.Category -TargetContext $TargetContext -ObjectDisplayName $row.ObjectDisplayName -ObjectIdentity $row.ObjectIdentity -ObjectAddress $row.ObjectAddress -ObjectType $row.ObjectType -Path $row.Path -Permission $row.Permission -OnPremisesSyncEnabled:$row.OnPremisesSyncEnabled -Removable:$row.Removable -Status $status -Reason $reason
    }

    return @($results)
}

function Transfer-DistributionOwnership {
    param(
        [Parameter(Mandatory)]
        $TargetContext,
        [Parameter(Mandatory)]
        [object[]]$Rows,
        [Parameter(Mandatory)]
        [string]$ReplacementIdentity
    )

    if (-not $TargetContext.ExchangeRecipient) {
        throw 'The target could not be resolved in Exchange Online.'
    }

    Ensure-ExchangeConnection
    $replacement = Get-Recipient -Identity $ReplacementIdentity -ErrorAction Stop
    $departingOwnerIdentity = if ($TargetContext.ExchangeRecipient.PrimarySmtpAddress) { [string]$TargetContext.ExchangeRecipient.PrimarySmtpAddress } else { $TargetContext.UserPrincipalName }
    $replacementKeySet = Resolve-KeySet -Set (New-KeySet)
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.PrimarySmtpAddress
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.DistinguishedName
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.Guid
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.ExternalDirectoryObjectId
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.UserPrincipalName
    Add-NormalizedValue -Set $replacementKeySet -Value $replacement.Alias

    $results = foreach ($row in $Rows) {
        $status = 'Skipped'
        $reason = $row.Reason

        if (-not $row.Transferable) {
            if ([string]::IsNullOrWhiteSpace($reason)) {
                $reason = 'This distribution list ownership entry is not transferable by this tool.'
            }
        }
        else {
            try {
                $group = Get-DistributionGroup -Identity $row.ObjectIdentity -ErrorAction Stop
                $replacementAlreadyOwner = $false
                foreach ($owner in @($group.ManagedBy)) {
                    if (Test-ObjectMatchesTarget -Object $owner -TargetKeys $replacementKeySet) {
                        $replacementAlreadyOwner = $true
                        break
                    }
                }

                $target = "{0} [{1}]" -f $row.ObjectDisplayName, $row.ObjectIdentity
                $action = "Transfer distribution list ownership from {0} to {1}" -f $departingOwnerIdentity, $replacement.PrimarySmtpAddress

                if ($script:ToolCmdlet.ShouldProcess($target, $action)) {
                    if (-not $replacementAlreadyOwner) {
                        Set-DistributionGroup -Identity $row.ObjectIdentity -ManagedBy @{Add=$replacement.PrimarySmtpAddress} -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                    }

                    Set-DistributionGroup -Identity $row.ObjectIdentity -ManagedBy @{Remove=$departingOwnerIdentity} -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                    $status = 'Transferred'
                    $reason = ''
                }
                else {
                    $reason = 'WhatIf or confirmation prevented the ownership transfer.'
                }
            }
            catch {
                $status = 'Error'
                $reason = $_.Exception.Message
            }
        }

        New-AuditRow -Operation 'TransferOwner' -Provider 'Exchange' -Category $row.Category -TargetContext $TargetContext -ObjectDisplayName $row.ObjectDisplayName -ObjectIdentity $row.ObjectIdentity -ObjectAddress $row.ObjectAddress -ObjectType $row.ObjectType -Path $row.Path -Permission $row.Permission -OnPremisesSyncEnabled:$row.OnPremisesSyncEnabled -Transferable:$row.Transferable -OwnerCount:$row.OwnerCount -Status $status -Reason $reason
    }

    return @($results)
}

function Remove-MailboxPermissions {
    param(
        [Parameter(Mandatory)]
        $TargetContext,
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if (-not $TargetContext.ExchangeRecipient) {
        throw 'The target could not be resolved in Exchange Online.'
    }

    Ensure-ExchangeConnection
    $delegateIdentifier = if ($TargetContext.ExchangeRecipient.PrimarySmtpAddress) { [string]$TargetContext.ExchangeRecipient.PrimarySmtpAddress } else { $TargetContext.UserPrincipalName }

    $results = foreach ($row in $Rows) {
        $status = 'Skipped'
        $reason = $row.Reason

        if (-not $row.Removable) {
            if ([string]::IsNullOrWhiteSpace($reason)) {
                $reason = 'This permission is not removable by this tool.'
            }
        }
        else {
            try {
                $target = "{0} [{1}]" -f $row.ObjectDisplayName, $row.ObjectIdentity
                $action = "Remove {0} {1} permission for {2}" -f $row.Permission, $row.ObjectType, $delegateIdentifier

                if ($script:ToolCmdlet.ShouldProcess($target, $action)) {
                    switch ($row.Permission) {
                        'FullAccess' {
                            Remove-MailboxPermission -Identity $row.ObjectIdentity -User $delegateIdentifier -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                        }
                        'SendAs' {
                            Remove-RecipientPermission -Identity $row.ObjectIdentity -Trustee $delegateIdentifier -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                        }
                        'SendOnBehalf' {
                            Set-Mailbox -Identity $row.ObjectIdentity -GrantSendOnBehalfTo @{Remove=$delegateIdentifier} -Confirm:$false -ErrorAction Stop
                        }
                        Default {
                            throw "Unsupported mailbox permission '$($row.Permission)'."
                        }
                    }

                    $status = 'Removed'
                    $reason = ''
                }
                else {
                    $reason = 'WhatIf or confirmation prevented removal.'
                }
            }
            catch {
                $status = 'Error'
                $reason = $_.Exception.Message
            }
        }

        New-AuditRow -Operation 'Remove' -Provider 'Exchange' -Category $row.Category -TargetContext $TargetContext -ObjectDisplayName $row.ObjectDisplayName -ObjectIdentity $row.ObjectIdentity -ObjectAddress $row.ObjectAddress -ObjectType $row.ObjectType -Path $row.Path -Permission $row.Permission -Removable:$row.Removable -Status $status -Reason $reason
    }

    return @($results)
}

function Show-Menu {
    param(
        [Parameter(Mandatory)]
        $TargetContext
    )

    Write-Host ''
    Write-Host 'Entra + Exchange Offboarding Access Tool' -ForegroundColor Green
    Write-Host '---------------------------------------' -ForegroundColor Green
    Write-Host ("Target user        : {0}" -f $TargetContext.UserPrincipalName)
    Write-Host ("Entra resolved     : {0}" -f [bool]$TargetContext.EntraUser)
    Write-Host ("Exchange resolved  : {0}" -f [bool]$TargetContext.ExchangeRecipient)
    Write-Host ("Removal enabled    : {0}" -f $script:EnableRemoval)
    Write-Host ("WhatIf             : {0}" -f [bool]$WhatIfPreference)
    Write-Host ("Log path           : {0}" -f $LogPath)
    if ($script:EnableRemoval) {
        Write-Host 'Mode               : Removal mode unlocked' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Mode               : Read-only (pass -EnableRemoval to unlock transfer/removal actions)' -ForegroundColor Cyan
    }
    Write-Host ''
    Write-Host 'Read-only actions' -ForegroundColor Cyan
    Write-Host '1. View Entra direct groups'
    Write-Host '2. View Entra effective groups (direct + transitive)'
    Write-Host '3. View distribution list membership and ownership'
    Write-Host '4. View mailbox delegate access (FullAccess, SendAs, SendOnBehalf)'
    Write-Host '5. View combined offboarding report'

    if ($script:EnableRemoval) {
        Write-Host ''
        Write-Host 'Transfer/removal actions' -ForegroundColor Yellow
        Write-Host '6. Remove selected Entra direct group memberships'
        Write-Host '7. Remove selected distribution list memberships'
        Write-Host '8. Transfer selected distribution list ownership to another user'
        Write-Host '9. Remove selected mailbox delegate access'
        Write-Host '10. Remove all removable Entra direct group memberships'
        Write-Host '11. Remove all removable distribution list memberships'
        Write-Host '12. Remove all removable mailbox delegate access'
        Write-Host '13. Change target user'
    }
    else {
        Write-Host '6. Change target user'
    }

    Write-Host 'Q. Quit'
    Write-Host ''
}

function Read-TargetUser {
    param(
        [string]$ExistingUpn
    )

    $upn = $ExistingUpn
    while ([string]::IsNullOrWhiteSpace($upn)) {
        $upn = Read-Host 'Enter the target user UPN'
    }

    return Resolve-TargetContext -UPN $upn
}

Ensure-LogPath -Path $LogPath
$targetContext = Read-TargetUser -ExistingUpn $UserPrincipalName

while ($true) {
    Show-Menu -TargetContext $targetContext
    $choice = (Read-Host 'Choose an option').Trim().ToUpperInvariant()

    switch ($choice) {
        '1' {
            $rows = Get-EntraGroupInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            Show-EntraInventory -Rows $rows -Title 'Entra direct groups'
        }
        '2' {
            $rows = Get-EntraGroupInventory -TargetContext $targetContext -IncludeTransitive
            Write-AuditRows -Rows $rows
            Show-EntraInventory -Rows $rows -Title 'Entra effective groups (direct + transitive)'
        }
        '3' {
            $rows = Get-DistributionListInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            Show-DistributionInventory -Rows $rows -Title 'Distribution list membership and ownership'
        }
        '4' {
            $rows = Get-MailboxPermissionInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            Show-MailboxInventory -Rows $rows -Title 'Mailbox delegate access'
        }
        '5' {
            $rows = Get-CombinedOffboardingReport -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            Show-CombinedInventory -Rows $rows -Title 'Combined offboarding report'
        }
        '6' {
            if ($script:EnableRemoval) {
                $rows = Get-EntraGroupInventory -TargetContext $targetContext
                Write-AuditRows -Rows $rows
                $eligible = @($rows | Where-Object { $_.Path -eq 'Direct' })
                $selected = Select-InventoryRows -Rows $eligible
                if ($selected.Count -gt 0) {
                    Write-Host ''
                    Write-Host 'Selected Entra groups' -ForegroundColor Yellow
                    Write-Host '---------------------' -ForegroundColor Yellow
                    $selected | Select-Object ObjectDisplayName, ObjectType, Removable, Reason | Format-Table -AutoSize
                    $confirm = Read-Host 'Type YES to continue, or anything else to cancel'
                    if ($confirm -eq 'YES') {
                        $results = Remove-EntraMemberships -TargetContext $targetContext -Rows $selected
                        Write-AuditRows -Rows $results
                        Show-EntraInventory -Rows $results -Title 'Entra removal results'
                    }
                    else {
                        Write-Host 'Removal canceled.' -ForegroundColor Yellow
                    }
                }
            }
            else {
                $newUpn = Read-Host 'Enter the new target user UPN'
                $targetContext = Read-TargetUser -ExistingUpn $newUpn
            }
        }
        '7' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-DistributionListInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $eligible = @($rows | Where-Object { $_.Category -eq 'DistributionMembership' })
            $selected = Select-InventoryRows -Rows $eligible
            if ($selected.Count -gt 0) {
                Write-Host ''
                Write-Host 'Selected distribution list memberships' -ForegroundColor Yellow
                Write-Host '--------------------------------------' -ForegroundColor Yellow
                $selected | Select-Object ObjectDisplayName, ObjectAddress, Removable, Reason | Format-Table -AutoSize
                $confirm = Read-Host 'Type YES to continue, or anything else to cancel'
                if ($confirm -eq 'YES') {
                    $results = Remove-DistributionMemberships -TargetContext $targetContext -Rows $selected
                    Write-AuditRows -Rows $results
                    Show-DistributionInventory -Rows $results -Title 'Distribution membership removal results'
                }
                else {
                    Write-Host 'Removal canceled.' -ForegroundColor Yellow
                }
            }
        }
        '8' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-DistributionListInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $eligible = @($rows | Where-Object { $_.Category -eq 'DistributionOwnership' })
            $selected = Select-InventoryRows -Rows $eligible -TransferOnly
            if ($selected.Count -gt 0) {
                Write-Host ''
                Write-Host 'Selected distribution list ownership entries' -ForegroundColor Yellow
                Write-Host '-------------------------------------------' -ForegroundColor Yellow
                $selected | Select-Object ObjectDisplayName, ObjectAddress, OwnerCount, Transferable, Reason | Format-Table -AutoSize
                $replacement = Read-Host 'Enter the replacement owner UPN or email address'
                if ([string]::IsNullOrWhiteSpace($replacement)) {
                    Write-Warning 'No replacement owner was provided. Transfer canceled.'
                    continue
                }

                $confirm = Read-Host 'Type TRANSFER to continue, or anything else to cancel'
                if ($confirm -eq 'TRANSFER') {
                    $results = Transfer-DistributionOwnership -TargetContext $targetContext -Rows $selected -ReplacementIdentity $replacement
                    Write-AuditRows -Rows $results
                    Show-DistributionInventory -Rows $results -Title 'Distribution ownership transfer results'
                }
                else {
                    Write-Host 'Transfer canceled.' -ForegroundColor Yellow
                }
            }
        }
        '9' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-MailboxPermissionInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $selected = Select-InventoryRows -Rows $rows
            if ($selected.Count -gt 0) {
                Write-Host ''
                Write-Host 'Selected mailbox delegate permissions' -ForegroundColor Yellow
                Write-Host '------------------------------------' -ForegroundColor Yellow
                $selected | Select-Object ObjectDisplayName, Permission, ObjectAddress, Removable, Reason | Format-Table -AutoSize
                $confirm = Read-Host 'Type YES to continue, or anything else to cancel'
                if ($confirm -eq 'YES') {
                    $results = Remove-MailboxPermissions -TargetContext $targetContext -Rows $selected
                    Write-AuditRows -Rows $results
                    Show-MailboxInventory -Rows $results -Title 'Mailbox permission removal results'
                }
                else {
                    Write-Host 'Removal canceled.' -ForegroundColor Yellow
                }
            }
        }
        '10' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-EntraGroupInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $selected = @($rows | Where-Object { $_.Path -eq 'Direct' -and $_.Removable })
            if (-not $selected -or $selected.Count -eq 0) {
                Write-Warning 'No removable Entra direct group memberships were found.'
                continue
            }

            Write-Host ''
            Write-Host 'Entra direct groups that will be targeted' -ForegroundColor Yellow
            Write-Host '---------------------------------------' -ForegroundColor Yellow
            $selected | Select-Object ObjectDisplayName, ObjectType | Format-Table -AutoSize
            $confirm = Read-Host 'Type REMOVEALL to continue, or anything else to cancel'
            if ($confirm -eq 'REMOVEALL') {
                $results = Remove-EntraMemberships -TargetContext $targetContext -Rows $selected
                Write-AuditRows -Rows $results
                Show-EntraInventory -Rows $results -Title 'Entra removal results'
            }
            else {
                Write-Host 'Bulk removal canceled.' -ForegroundColor Yellow
            }
        }
        '11' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-DistributionListInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $selected = @($rows | Where-Object { $_.Category -eq 'DistributionMembership' -and $_.Removable })
            if (-not $selected -or $selected.Count -eq 0) {
                Write-Warning 'No removable distribution list memberships were found.'
                continue
            }

            Write-Host ''
            Write-Host 'Distribution list memberships that will be targeted' -ForegroundColor Yellow
            Write-Host '---------------------------------------------------' -ForegroundColor Yellow
            $selected | Select-Object ObjectDisplayName, ObjectAddress | Format-Table -AutoSize
            $confirm = Read-Host 'Type REMOVEALL to continue, or anything else to cancel'
            if ($confirm -eq 'REMOVEALL') {
                $results = Remove-DistributionMemberships -TargetContext $targetContext -Rows $selected
                Write-AuditRows -Rows $results
                Show-DistributionInventory -Rows $results -Title 'Distribution membership removal results'
            }
            else {
                Write-Host 'Bulk removal canceled.' -ForegroundColor Yellow
            }
        }
        '12' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Removal mode is disabled. Re-run with -EnableRemoval to unlock destructive actions.'
                continue
            }

            $rows = Get-MailboxPermissionInventory -TargetContext $targetContext
            Write-AuditRows -Rows $rows
            $selected = @($rows | Where-Object { $_.Removable })
            if (-not $selected -or $selected.Count -eq 0) {
                Write-Warning 'No removable mailbox delegate permissions were found.'
                continue
            }

            Write-Host ''
            Write-Host 'Mailbox permissions that will be targeted' -ForegroundColor Yellow
            Write-Host '----------------------------------------' -ForegroundColor Yellow
            $selected | Select-Object ObjectDisplayName, Permission, ObjectAddress | Format-Table -AutoSize
            $confirm = Read-Host 'Type REMOVEALL to continue, or anything else to cancel'
            if ($confirm -eq 'REMOVEALL') {
                $results = Remove-MailboxPermissions -TargetContext $targetContext -Rows $selected
                Write-AuditRows -Rows $results
                Show-MailboxInventory -Rows $results -Title 'Mailbox permission removal results'
            }
            else {
                Write-Host 'Bulk removal canceled.' -ForegroundColor Yellow
            }
        }
        '13' {
            if (-not $script:EnableRemoval) {
                Write-Warning 'Invalid selection.'
                continue
            }

            $newUpn = Read-Host 'Enter the new target user UPN'
            $targetContext = Read-TargetUser -ExistingUpn $newUpn
        }
        'Q' {
            break
        }
        Default {
            Write-Warning 'Invalid selection.'
        }
    }
}
