
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"

# 1. Get all users with required properties
$allUsers = Get-MgUser -All -Property "id,displayName,mail,userPrincipalName,assignedLicenses,accountEnabled"

# 2. Filter to licensed users with sign-in blocked
$blockedLicensedUsers = $allUsers | Where-Object {
    $_.AccountEnabled -eq $false -and
    $_.AssignedLicenses.Count -gt 0
}

Write-Host "===== BLOCKED LICENSED USERS (RAW UPNs) =====" -ForegroundColor Cyan

# ---- LIST 1: Raw UPNs + some basic info in the console ----
$blockedLicensedUsers |
    Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled |
    Format-Table -AutoSize

Write-Host "`n===== BLOCKED LICENSED USERS - UPN LIST (RAW) =====" -ForegroundColor Yellow

# Just the current UPNs, one per line in the console
$blockedLicensedUsers |
    Select-Object -ExpandProperty UserPrincipalName

# Export raw UPNs to CSV (single column)
$blockedLicensedUsers |
    Select-Object UserPrincipalName |
    Export-Csv -Path ".\Blocked_Licensed_UPNs_Raw.csv" -NoTypeInformation

Write-Host "`nRaw UPN list exported to .\Blocked_Licensed_UPNs_Raw.csv" -ForegroundColor Green


# ---- LIST 2: UPNs in first.last@quantinuum.com format ----

Write-Host "`n===== BLOCKED LICENSED USERS - UPN LIST (first.last@quantinuum.com) =====" -ForegroundColor Cyan

# Build transformed UPNs from DisplayName
$blockedLicensedTransformed = $blockedLicensedUsers | ForEach-Object {
    # Split DisplayName on spaces (assumes 'First Last' or 'First Middle Last')
    $nameParts = ($_.DisplayName -split '\s+')
    $first = $nameParts[0]
    $last  = $nameParts[-1]

    # Construct first.last@quantinuum.com
    $newUpn = ("{0}.{1}@quantinuum.com" -f $first, $last).ToLower()

    # Output a custom object with both original and transformed UPN if you want it
    [PSCustomObject]@{
        DisplayName          = $_.DisplayName
        OriginalUserPrincipalName = $_.UserPrincipalName
        TransformedUPN       = $newUpn
    }
}

# Show the transformed list in the console
$blockedLicensedTransformed |
    Select-Object DisplayName, OriginalUserPrincipalName, TransformedUPN |
    Format-Table -AutoSize

Write-Host "`n===== TRANSFORMED UPNs ONLY (first.last@quantinuum.com) =====" -ForegroundColor Yellow

$blockedLicensedTransformed |
    Select-Object -ExpandProperty TransformedUPN

# Export only the transformed UPNs to CSV
$blockedLicensedTransformed |
    Select-Object TransformedUPN |
    Export-Csv -Path ".\Blocked_Licensed_UPNs_Transformed.csv" -NoTypeInformation

Write-Host "`nTransformed UPN list exported to .\Blocked_Licensed_UPNs_Transformed.csv" -ForegroundColor Green