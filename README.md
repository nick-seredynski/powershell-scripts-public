# Powershell scripts for Microsoft environments

Most of the scripts require Cloud based infastructure through Azure but may also work for Hybrid environments.
Scripts can be used for automating tasks such as automatically onboarding and offboarding users, creating backups and monitoring infastructure. For most scripts I included automated updates and error handling as seen below:

---
For example function "Check-InstallModule" will check if modules are installed and up to date, if not installed script will install or update them automatically.

<img width="657" height="359" alt="image" src="https://github.com/user-attachments/assets/bd36c8b2-029d-4ac5-9f8f-1e1f8a2fcb8d" />

---
I included error handling to return errors in case any part of the script fails

<img width="691" height="163" alt="image" src="https://github.com/user-attachments/assets/d176e30b-0e2e-46d3-ba16-405897bc86ff" />


## What each script does:

**NewUser.ps1 :** Creates a new user by copying an existing user, this includes groups and permissions. The license will then need to be assigned manually after the script is finished running.

**OffboardUser.ps1 :** Will offboard a user by blocking the account, adding suffix (shared), chaging mailbox type to shared mailbox, giving manager or another user access to mailbox and creating autoreplies to inform whoever emails the inbox to person in question. Finally the account is changed to a guest type.

**RemoveGroups.ps1 :** Simply strips away all of a user's groups on entra.

**ShareCalendar.ps1 :** Gives a GranteeUser access to TargetUser's calendar, as an admin this is the only way to do it without directly accessing the mailbox.

**CopyUser.ps1 :** Copies permissions from another user, can be used after NewUser.ps1 to ensure permissions are synced correctly

---
## How to run scripts:
1. The entire script can be copy and pasted into powershell.
2. Save the script somewhere as a .ps1 file and then right click 'run in powershell'
3. Paste the following into powershell: powershell -noexit "& ""C:\Users\NickSeredynski\Scripts\ExampleScript.ps1"""  make sure the fromat follows the structure of the folders whhere the tagret script is located, in this case the last folder is 'Scripts' and name of the script is 'Example script'.
```
powershell -noexit "& ""C:\Users\NickSeredynski\Scripts\ExampleScript.ps1"""
```
---
## Tips and troubleshooting:
In case of this error message:
```
In case of this error ...is not digitally signed. You cannot run this script on the current system.    + CategoryInfo          : SecurityError: (:) [], PSSecurityException
    + FullyQualifiedErrorId : UnauthorizedAccess
```
By default a script may be blocked to prevent unathorised scripts for being run, you may choose to unblock it by pointing powershell to the location of the script as seen in this example:
```
Unblock-File -Path "C:\Users\NickSeredynski\Scripts\ScriptName.ps1"
```
Alternatively, not recommended, you can choose to Set the sexecution policy to remote signed which unblocks all scripts used in the organisation:
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
```
And then  this to once again block running scripts in powershell:
```
Set-ExecutionPolicy Restricted
```

