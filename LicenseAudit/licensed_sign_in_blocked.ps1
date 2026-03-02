
# Licensed users with sign-in blocked + list of UPNs
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"

# 1. Get all users with the needed properties
$allUsers = Get-MgUser -All -Property "id,displayName,mail,userPrincipalName,assignedLicenses,accountEnabled"

# 2. Filter to licensed users whose sign-in is blocked
#    - accountEnabled -eq $false  -> sign-in blocked
#    - assignedLicenses.Count -gt 0 -> has at least one license assigned
$blockedLicensedUsers = $allUsers | Where-Object {
    $_.AccountEnabled -eq $false -and
    $_.AssignedLicenses.Count -gt 0
}

Write-Host "Blocked, licensed users:" -ForegroundColor Cyan

# 3. Show a nice table in the PowerShell console
$blockedLicensedUsers |
    Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled |
    Format-Table -AutoSize

# 4. Build a list of UPNs in first.last@quantinuum.com format
#    - If your existing UPNs are already in that form, you can just use UserPrincipalName directly.
#    - If not, this example converts from DisplayName -> 'first.last@quantinuum.com'.

$blockedLicensedUpns = $blockedLicensedUsers | ForEach-Object {
    # Try to derive first.last from DisplayName if needed
    # Assumes 'First Last' style names
    $nameParts = ($_.DisplayName -split '\s+') # split on space
    $first = $nameParts[0]
    $last  = $nameParts[-1]

    # Build the UPN in the desired format
    ("{0}.{1}@quantinuum.com" -f $first, $last).ToLower()
}

# 5. Output the UPN list to the console
Write-Host "`nBlocked, licensed UPNs (first.last@quantinuum.com):" -ForegroundColor Yellow
$blockedLicensedUpns | ForEach-Object { Write-Host $_ }

# 6. Export just the UPNs to CSV (single column)
$blockedLicensedUpns |
    ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_ } } |
    Export-Csv -Path ".\Blocked_Licensed_UPNs.csv" -NoTypeInformation

Write-Host "`nCSV exported: .\Blocked_Licensed_UPNs.csv" -ForegroundColor Green