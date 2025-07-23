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
Check-InstallModule -ModuleName "ExchangeOnlineManagement"


# Import necessary modules
Import-Module AzureAD
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

# Connect to Entra ID 
Connect-AzureAD

# Prompt for admin UPN and connect to Exchange Online
$AdminUPN = Read-Host "Enter your admin email address (UPN)"
$Session = Connect-ExchangeOnline -UserPrincipalName $AdminUPN


# Define the user account to be locked and converted
$UserToLock = Read-Host "Enter username to lock"

# Define the users who will get read/write permissions
$manager_email = Read-Host "Enter user's manager's email"


# Lock the user account
# Disable the user account (Microsoft Graph)
try {
    Update-MgUser -UserId $UserToLock -AccountEnabled:$false
    Write-Host "User account $UserToLock disabled successfully via Microsoft Graph."
} catch {
    Write-Host "Error disabling user account via Microsoft Graph: $_"
}

# Disable via AzureAD module for compatibility
try {
    Set-AzureADUser -ObjectId $UserToLock -AccountEnabled $false
    Write-Host "User account $UserToLock disabled successfully via AzureAD."
} catch {
    Write-Host "Error disabling user account via AzureAD: $_"
}


# Convert the user mailbox to a shared mailbox
try {
    $UserMailbox = Get-Mailbox -Identity $UserToLock
    if ($UserMailbox) {
        Set-Mailbox -Identity $UserMailbox -Type Shared
        Write-Host "User mailbox converted to shared mailbox successfully."
    } else {
        Write-Host "User mailbox not found."
    }
} catch {
    Write-Host "Error converting mailbox: $_"
}


# Grant read/write permissions for the user's email inbox to the manager
try {
    $manager_upn = (Get-User -Identity $manager_email).UserPrincipalName
    if ($manager_upn) {
        Add-MailboxPermission -Identity $UserToLock -User $manager_upn -AccessRights FullAccess -AutoMapping:$false
        Write-Host "Permissions granted to $manager_upn successfully."
    } else {
        Write-Host "Manager UPN not found for $manager_email"
    }
} catch {
    Write-Host "Error granting permissions: $_"
}


# Set up automatic replies

$AutoReplyMessage = "Thank you for your email, I am currently out of the office with no access to emails - please contact $manager_email for assistance. This is an automated reply. For your convenience, this email has been automatically forwarded to $manager_email via a shared mailbox."

try {
    if ($UserMailbox) {
        Set-MailboxAutoReplyConfiguration -Identity $UserMailbox -AutoReplyState Enabled -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage -ExternalAudience All
        Write-Host "Automatic replies enabled successfully."
    } else {
        Write-Host "Cannot set up automatic replies because the user mailbox was not found."
    }
} catch {
    Write-Host "Error setting up automatic replies: $_"
}


# Append (Shared) to name, eg "John Doe" > "John Doe (Shared)" 
try {
    # Get the user object
    $user = Get-MgUser -UserId $UserToLock

    if ($user -ne $null) {
        # Avoid appending "(Shared)" multiple times
        if ($user.DisplayName -notlike "*(Shared)*") {
            $newDisplayName = "$($user.DisplayName) (Shared)"
            Update-MgUser -UserId $UserToLock -DisplayName $newDisplayName
            Write-Host "Display name updated to '$newDisplayName' for user $UserToLock."
        } else {
            Write-Host "Display name already contains '(Shared)'. No update needed."
        }
    } else {
        Write-Warning "User not found: $UserToLock"
    }
}
catch {
    Write-Error "Error updating display name for user ${UserToLock}: $_"
}


# Attempt to change user type from Member to Guest

try {
    # Get the user object first
    $user = Get-MgUser -UserId $UserToLock

    if ($user -ne $null) {
        # Convert the user to a guest
        Update-MgUser -UserId $UserToLock -UserType Guest

        Write-Host "User '$UserToLock' account converted to guest."
    } else {
        Write-Warning "User not found: $UserToLock"
    }
}
catch {
    Write-Error "Error converting user account to guest: $_"
}
