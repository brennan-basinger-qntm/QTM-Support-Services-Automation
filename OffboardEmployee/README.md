# Offboarding Toolkit (Rebuilt)

This is a hardened PowerShell automation for employee offboarding:
**snapshot -> plan -> apply -> snapshot -> ServiceNow notes**.

## Quick start
1. Open **PowerShell 7 (x64)** as Administrator.
2. Run once per machine:
   ```powershell
   .\src\Bootstrap-Env.ps1
   ```
3. Dry‑run (no changes by default):
   ```powershell
   .\src\Invoke-UserOffboarding.ps1 -UserUpn 'first.last@quantinuum.com' -TicketNumber 'INC12345678'
   ```
4. When the plan looks right, add `-Apply` to perform changes.

## Common switches
- `-SupervisorUpn 'manager.name@quantinuum.com' -GrantSupervisorFullAccess -GrantSupervisorSendAs`
- `-RemoveLicenses -DisableEntraSignIn`
- `-DisableAD -UpdateAdDescription -DisabledOuDn "OU=Disabled,OU=Corp,DC=yourco,DC=com"`

## What it does
- Uses **Graph** for Entra groups/licensing and **EXO** for DLs, mailbox, and delegations.
- Skips **dynamic** groups (lists them for visibility).
- Converts mailbox to **Shared** and stamps an **expiry** marker in `CustomAttribute15`.
- Creates **Before / After** CSVs and **ServiceNow work notes** you can paste into the ticket.
- Runs in **Preview** unless `-Apply` is present.

## Requirements
- PowerShell 7.4+ (x64)
- Modules (script auto-installs if missing): `ExchangeOnlineManagement` (≥3.3), `Microsoft.Graph` (≥2.16)

## Safety
- No changes occur without `-Apply`.
- Every run writes a transcript and CSVs into a timestamped folder on that's output to Desktop by default.

**Support Services ready.**
