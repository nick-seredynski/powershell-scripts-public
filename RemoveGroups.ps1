# Function to check if a module is installed, and install it if not
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

# Prompt for admin UPN and connect to Exchange Online
$AdminUPN = Read-Host "Enter your admin email address (UPN)"
$Session = Connect-ExchangeOnline -UserPrincipalName $AdminUPN

Check-InstallModule -ModuleName "ExchangeOnlineManagement"

# Import necessary modules
Import-Module AzureAD
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users

# Connect to Entra ID and Exchange Online
Connect-AzureAD
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

# Optional: Confirm the connection and see the current context
$context = Get-MgContext

# Define the user account to be locked and converted
$UserToLock = Read-Host "Enter the user's email to lock"

if (-not $UserToLock) {
    Write-Error "User email not provided."
    return
}


# Get the user's object ID from Microsoft Graph
$user = Get-MgUser -UserId $UserToLock

if (-not $user) {
    Write-Error "User $UserToLock not found in Microsoft Graph."
    return
}


# Confirm before removing from all groups
$confirmation = Read-Host "Do you want to remove $UserToLock from all Entra ID groups? (Y/N)"
if ($confirmation -eq "Y") {
    Write-Host "Fetching groups for $UserToLock..."

    # Get all groups the user is a member of (paged)
    $groups = @()
    $response = Get-MgUserMemberOf -UserId $user.Id -All

    foreach ($item in $response) {
        if ($item.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
            $groups += $item
        }
    }

    if ($groups.Count -eq 0) {
        Write-Host "$UserToLock is not a member of any groups."
    } else {
        Write-Host "$UserToLock is a member of $($groups.Count) groups. Proceeding with removal..."

        foreach ($group in $groups) {
            try {
                $groupId = $group.Id
                $groupName = $group.AdditionalProperties["displayName"]

                Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $user.Id
                Write-Host " Removed $UserToLock from group: $groupName"
            } catch {
                Write-Warning " Failed to remove $UserToLock from group: $groupName - $_"
            }
        }
    }
} else {
    Write-Host "Skipping group removal."
}

# Disconnect from session
Disconnect-ExchangeOnline -Confirm:$false at the end.

