param (
    [Parameter(Mandatory = $true)]
    [string]$TargetUser,

    [Parameter(Mandatory = $true)]
    [string]$GranteeUser,

    [Parameter(Mandatory = $false)]
    [ValidateSet("None", "AvailabilityOnly", "LimitedDetails", "Reviewer", "Contributor", "NonEditingAuthor", "Author", "PublishingAuthor", "Editor", "PublishingEditor", "Owner", "Custom")]
    $PermissionLevel = Read-Host "Enter the permission level (e.g., Reviewer, Editor, Owner). Press Enter for default (Editor)"
    if ([string]::IsNullOrWhiteSpace($PermissionLevel)) {
        $PermissionLevel = "Editor"
}

)

# Install the Exchange Online module if not already installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}

# Import the module
Import-Module ExchangeOnlineManagement

# Prompt for admin UPN and connect to exchange online
$AdminUPN = Read-Host "Enter your admin email address (UPN)"
$Session = Connect-ExchangeOnline -UserPrincipalName $AdminUPN

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
