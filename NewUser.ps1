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

# Connect to Azure AD and Exchange Online via admin
Connect-AzureAD
$AdminUPN = Read-Host "Enter your admin email address (UPN)"
$Session = Connect-ExchangeOnline -UserPrincipalName $AdminUPN

# Step 1: Prompt for New User Details and Create the User
$firstName = Read-Host "Enter the first name of the new user"
$lastName = Read-Host "Enter the last name of the new user"
$email = Read-Host "Enter the email address of the new user"
$password = Read-Host "Enter a temporary password for the new user" -AsSecureString

Write-Host "Creating new user $firstName $lastName with email $email..."

$newUser = New-AzureADUser -DisplayName "$firstName $lastName" `
    -GivenName $firstName `
    -Surname $lastName `
    -UserPrincipalName $email `
    -MailNickname ($email.Split("@")[0]) `
    -AccountEnabled $true `
    -PasswordProfile @{Password=$password; ForceChangePasswordNextLogin=$true}

if (!$newUser) {
    Write-Error "Failed to create new user. Aborting process."
    return
}

Write-Host "New user $email created successfully."

# Step 2: Specify the Source User
$sourceUserUPN = Read-Host "Enter the UPN of the source user to copy permissions from"

$sourceUser = Get-AzureADUser -ObjectId $sourceUserUPN
if (!$sourceUser) {
    Write-Error "Source user $sourceUserUPN not found. Aborting process."
    return
}

Write-Host "Source User: $sourceUser.DisplayName"

# Step 3: Copy Group Memberships (Exclude Administrative Roles)
Write-Host "Copying group memberships from $sourceUserUPN to $email..."
$groups = Get-AzureADUserMembership -ObjectId $sourceUser.ObjectId | Where-Object {
    $_.ODataType -notlike "*DirectoryRole*" -and $_.DisplayName -notmatch "(Administrator|Admin|Role)"
}

foreach ($group in $groups) {
    try {
        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $newUser.ObjectId
        Write-Host "Added $email to group $($group.DisplayName)."
    } catch {
        Write-Warning "Could not add $email to group $($group.DisplayName): $_"
    }
}

# Pause for 3 minutes and prompt user to assign license manually
Write-Host "Please assign the user a 365 Business Premium license manually in order to generate a mailbox for the user."
Start-Sleep -Seconds 120
Write-Host "Resuming script..."

# Step 4: Copy Shared Mailbox Permissions
Write-Host "Copying shared mailbox permissions..."
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | Where-Object {
    (Get-MailboxPermission -Identity $_.PrimarySmtpAddress | Where-Object { $_.User -eq $sourceUserUPN })
}

foreach ($sharedMailbox in $sharedMailboxes) {
    $permissions = Get-MailboxPermission -Identity $sharedMailbox.PrimarySmtpAddress | Where-Object {
        $_.User -eq $sourceUserUPN
    }

    foreach ($permission in $permissions) {
        try {
            Add-MailboxPermission -Identity $sharedMailbox.PrimarySmtpAddress `
                -User $email `
                -AccessRights $permission.AccessRights `
                -InheritanceType $permission.InheritanceType
            Write-Host "Copied permissions for shared mailbox $($sharedMailbox.PrimarySmtpAddress)."
        } catch {
            Write-Warning "Could not copy permissions for shared mailbox $($sharedMailbox.PrimarySmtpAddress): $_"
        }
    }
}

# Step 5: Copy Custom Attributes (Including Company Name and Manager)
Write-Host "Copying custom attributes..."
try {
    Set-AzureADUser -ObjectId $newUser.ObjectId `
        -Department $sourceUser.Department `
        -JobTitle $sourceUser.JobTitle `
        -City $sourceUser.City `
        -Country $sourceUser.Country `
        -CompanyName $sourceUser.CompanyName
    Write-Host "Copied custom attributes successfully."
} catch {
    Write-Warning "Could not copy custom attributes: $_"
}

# Copy Manager
Write-Host "Copying manager information..."
try {
    $manager = Get-AzureADUserManager -ObjectId $sourceUser.ObjectId
    if ($manager) {
        Set-AzureADUserManager -ObjectId $newUser.ObjectId -RefObjectId $manager.ObjectId
        Write-Host "Copied manager: $($manager.DisplayName)"
    } else {
        Write-Host "Source user has no manager assigned."
    }
} catch {
    Write-Warning "Could not copy manager: $_"
}

# Step 6: Copy Distribution List Memberships
Write-Host "Copying distribution list memberships..."
$distributionLists = Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup | Where-Object {
    (Get-DistributionGroupMember -Identity $_.PrimarySmtpAddress | Where-Object { $_.PrimarySmtpAddress -eq $sourceUserUPN })
}

foreach ($dl in $distributionLists) {
    try {
        Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $email
        Write-Host "Added $email to distribution list $($dl.DisplayName)."
    } catch {
        Write-Warning "Could not add $email to distribution list $($dl.DisplayName): $_"
    }
}

Write-Host "Permissions copying process completed successfully."
