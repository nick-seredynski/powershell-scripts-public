# Powershell scripts for Microsoft environments

---
## Unblocking a script:
```
In case of this error ...is not digitally signed. You cannot run this script on the current system.    + CategoryInfo          : SecurityError: (:) [], PSSecurityException
    + FullyQualifiedErrorId : UnauthorizedAccess
```
To unblock a specific file:
```
Unblock-File -Path "C:\Users\NickSeredynski\Scripts\ScriptName.ps1"
```
To set execution policy for all files
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
```
---
## Example of how to run a script:
In this example the script "NewUser.ps1" is saved in the following directory C:\Users\NickSeredynski\Scripts\
```
powershell -noexit "& ""C:\Users\NickSeredynski\Scripts\NewUser.ps1"""
```
