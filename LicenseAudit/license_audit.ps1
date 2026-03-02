# E3 and Defender License Audit (UPNs: first.last or f.last)

Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"

$E3SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"                 # E3
$DefenderBundleSkuId = "3dd6cf57-d688-4eed-ba52-9e40b5468c3e"     # Defender

# Match:
# - first.last@quantinuum.com
# - f.last@quantinuum.com
# Only letters on both sides of the dot, exactly one dot, quantinuum.com domain
$PersonUpnPattern = '^[A-Za-z]{1,}\.[A-Za-z]{1,}@quantinuum\.com$'

# Pull users (add consistency flags for large directories)
$allUsers = Get-MgUser -All `
    -Property "id,displayName,mail,userPrincipalName,assignedLicenses" `
    -ConsistencyLevel eventual `
    -CountVariable totalUserCount

# Helper: keep only UPNs that match the person format
function Test-PersonUpn {
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $false }
    return ($Upn -match $PersonUpnPattern)
}

# Users with E3 but WITHOUT the Defender bundle, and UPN matches person format
$E3_No_Defender = $allUsers | Where-Object {
    ($_.assignedLicenses.skuId -contains [guid]$E3SkuId) -and
    -not ($_.assignedLicenses.skuId -contains [guid]$DefenderBundleSkuId) -and
    (Test-PersonUpn $_.UserPrincipalName)
}

# Users with the Defender bundle but WITHOUT E3, and UPN matches person format
$Defender_No_E3 = $allUsers | Where-Object {
    ($_.assignedLicenses.skuId -contains [guid]$DefenderBundleSkuId) -and
    -not ($_.assignedLicenses.skuId -contains [guid]$E3SkuId) -and
    (Test-PersonUpn $_.UserPrincipalName)
}

# Export as a clean list (no header, no column name)
$E3_No_Defender |
    Select-Object -ExpandProperty UserPrincipalName |
    Sort-Object -Unique |
    Set-Content -Path ".\E3_without_DefenderP2_upn_person_only.txt" -Encoding UTF8

$Defender_No_E3 |
    Select-Object -ExpandProperty UserPrincipalName |
    Sort-Object -Unique |
    Set-Content -Path ".\DefenderP2_without_E3_upn_person_only.txt" -Encoding UTF8

# Optional: also export CSV versions (keeps a header)
$E3_No_Defender |
    Select-Object UserPrincipalName |
    Sort-Object UserPrincipalName -Unique |
    Export-Csv -Path ".\E3_without_DefenderP2_upn_person_only.csv" -NoTypeInformation

$Defender_No_E3 |
    Select-Object UserPrincipalName |
    Sort-Object UserPrincipalName -Unique |
    Export-Csv -Path ".\DefenderP2_without_E3_upn_person_only.csv" -NoTypeInformation
``