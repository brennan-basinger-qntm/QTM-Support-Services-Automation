# Employee Offboarding Script (Step by Step Guide)

This folder contains a PowerShell script that helps the Information Technology team offboard an employee in Microsoft 365.

The script is designed to be safe.
- It runs in preview mode by default, which means it shows what it would do, but it does not change anything.
- It creates a before snapshot and an after snapshot.
- It writes ServiceNow work notes you can paste into the ticket.

The offboarding script is named `Invoke-UserOffboarding.ps1`.

## What the script does

When you run the script in apply mode, it can do these actions:
- Convert the user mailbox to a shared mailbox and write an expiry reminder into CustomAttribute15.
- Remove the user from static distribution groups and mail enabled security groups.
- Remove the user from static Microsoft 365 groups and security groups.
- Remove mailbox delegation on the user mailbox (Full Access, Send As, Send on Behalf) that other people have.
- Remove the offboarded user access to other mailboxes (for example, shared mailboxes they had access to).
- Remove Microsoft 365 licenses from the user.
- Block sign in for the user and revoke sign in sessions.
- Optional: perform on premises Active Directory steps if requested and available.

The script never changes dynamic group membership. Dynamic groups are rule based and are listed for visibility only.

## Before you start

### You need
1. A Windows computer.
2. PowerShell version 7 or newer, 64 bit.
3. Internet access.
4. Permission to manage users, groups, licenses, and mailboxes in our Microsoft 365 environment.
5. The script files downloaded to your computer.

### One time setup on each computer

Run the bootstrap script one time per computer. This installs the PowerShell modules the offboarding script needs.

1. Open the Start menu.
2. Type **PowerShell 7**.
3. Right click **PowerShell 7** and click **Run as administrator**.
4. In the PowerShell window, run this command so the window can run scripts:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```

5. Go to the folder that contains the scripts. Example:

```powershell
cd "C:\Path\To\OffboardEmployee\src"
```

6. Run the bootstrap script:

```powershell
.\Bootstrap-Env.ps1
```

If you see messages about installing modules, let it finish.

## Offboarding a user (preview first)

### Step 1. Get the required information

You need:
- The user email address (example: `first.last@quantinuum.com`).
- The ServiceNow ticket number (example: `INC12345678` or `SCTASK1234567`).

Optional, but often useful:
- Supervisor email address, if the supervisor needs mailbox access.
- Backup owner email address, if the user owns Microsoft 365 groups.

### Step 2. Run the script in preview mode

Preview mode is the default. Do not add `-Apply`.

```powershell
.\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC12345678"
```

### Step 3. Review the output folder

At the end of the run, the script prints an evidence folder path. By default, it creates a new folder on your Desktop.

Inside that folder you will see:
- A plan file that lists what the script plans to change.
- Before snapshot files.
- After snapshot files.
- A transcript log file.
- A text file named `ServiceNow-WorkNotes.txt`.

Open the plan file and make sure the plan matches what the ticket requests.

## Apply changes (only after you reviewed the plan)

When you are ready to make changes, run the same command again but add `-Apply`.

```powershell
.\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC12345678" -Apply
```

After it finishes, open the evidence folder again and review the after snapshot files.

## Common options you may need

### Give the supervisor access to the mailbox

This grants the supervisor Full Access and Send As.

```powershell
.\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC12345678" -SupervisorUpn "manager.name@quantinuum.com" -GrantSupervisorFullAccess -GrantSupervisorSendAs
```

To apply changes, add `-Apply`.

### Add a backup owner to groups the user owns

If the user is the only owner of a Microsoft 365 group, the script can add a backup owner.

```powershell
.\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC12345678" -BackupOwnerUpn "backup.owner@quantinuum.com"
```

To apply changes, add `-Apply`.

### Optional on premises Active Directory steps

These steps only run if the Active Directory module is available on the computer.

Example:

```powershell
.\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC12345678" -DisableAD -UpdateAdDescription -Apply
```

## ServiceNow work notes

After a run, open `ServiceNow-WorkNotes.txt` in the evidence folder and paste it into the ticket work notes.

## Troubleshooting

### The shared mailbox does not show up for the supervisor

This can happen even after permissions are granted.
Try these steps:
- Wait a few minutes.
- Restart Outlook.
- If it still does not show, add the mailbox manually in Outlook account settings.

### Apps still work after sign in is blocked

Blocking sign in stops new sign ins, but existing sessions can continue.
Make sure session revocation ran successfully, then test again.

### Module install problems

If module installation fails, close PowerShell, re open PowerShell 7 as administrator, and re run `Bootstrap-Env.ps1`.

## Safety rules

1. Always run preview first.
2. Read the plan file before applying.
3. Do not offboard the wrong person. Double check the user email address.
4. Keep the evidence folder. It is your proof of what happened.
