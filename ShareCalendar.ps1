function Check-InstallModule {
    param (
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module $ModuleName not found. Installing..."
        Install-Module -Name $ModuleName -Force -AllowClobber
    } else {
        Write-Host "Module $ModuleName is already installed. Skipping installation."
    }
}

# Check and install required modules if necessary
Check-InstallModule -ModuleName "AzureAD"
Check-InstallModule -ModuleName "ExchangeOnlineManagement"

# Import necessary modules
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Check and install required modules if necessary
Check-InstallModule -ModuleName "ExchangeOnlineManagement"

# Prompt for admin UPN and connect to exchange online
$AdminUPN = Read-Host "Enter your admin email address (UPN)"
Connect-ExchangeOnline -UserPrincipalName $AdminUPN

# Prompt for inputs
$TargetUser = Read-Host "Enter the user's UPN who's mailbox will be shared (TargetUser)"
$GranteeUser = Read-Host "Enter the user UPN who will be granted access (GranteeUser)"

# If field is empty default permission will be editor
$PermissionLevel = Read-Host "Enter the permission level (Default = Editor). Options: None, AvailabilityOnly, LimitedDetails, Reviewer, Contributor, NonEditingAuthor, Author, PublishingAuthor, Editor, PublishingEditor, Owner, Custom"

if ([string]::IsNullOrWhiteSpace($PermissionLevel)) {
    $PermissionLevel = "Editor"
}

# Check if permission exists
try {
    $Permission = Get-MailboxFolderPermission -Identity "$($TargetUser):\Calendar" -User $GranteeUser -ErrorAction Stop
    # If permission exists, update it
    Set-MailboxFolderPermission -Identity "$($TargetUser):\Calendar" -User $GranteeUser -AccessRights $PermissionLevel
    Write-Host "Existing permission updated successfully!"
} catch {
    # If permission doesn't exist, add a new one
    Add-MailboxFolderPermission -Identity "$($TargetUser):\Calendar" -User $GranteeUser -AccessRights $PermissionLevel
    Write-Host "New permission entry added successfully!"
}


# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Calendar access granted successfully!"
