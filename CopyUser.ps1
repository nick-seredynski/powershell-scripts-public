# Import necessary modules and connect to services
Import-Module AzureAD
Connect-AzureAD

# Variables
$sourceUserUPN = Read-Host "UPN/Email of the user where permissions are copied"  # User who's permissions are being copied from
$targetUserUPN = Read-Host "UPN/Email of the user where permissions are applied"  # User who's permissions are copied to

# Step 1: Fetch User Details
$sourceUser = Get-AzureADUser -ObjectId $sourceUserUPN
$targetUser = Get-AzureADUser -ObjectId $targetUserUPN

if (!$sourceUser) {
    Write-Error "Source user $sourceUserUPN not found."
    return
}

if (!$targetUser) {
    Write-Error "Target user $targetUserUPN not found."
    return
}

Write-Host "Source User: $sourceUser.DisplayName"
Write-Host "Target User: $targetUser.DisplayName"

# Step 2: Copy Group Memberships (Exclude Administrative Roles)
Write-Host "Copying group memberships (excluding roles like Compliance Administrator)..."

# Fetch all group memberships but exclude Azure AD roles and admin-related groups
$groups = Get-AzureADUserMembership -ObjectId $sourceUser.ObjectId | Where-Object {
    $_.ODataType -notlike "*DirectoryRole*" -and $_.DisplayName -notmatch "(Administrator|Admin|Role)"
}

foreach ($group in $groups) {
    try {
        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $targetUser.ObjectId
        Write-Host "Added $targetUserUPN to group $($group.DisplayName)."
    } catch {
        if ($_.Exception.Message -like "*ResourceNotFound*") {
            Write-Warning "Group $($group.DisplayName) does not exist or cannot be accessed. Skipping."
        } else {
            Write-Warning "Could not add $targetUserUPN to group $($group.DisplayName): $_"
        }
    }
}

# Step 3: Copy Shared Mailbox Permissions
Write-Host "Copying shared mailbox permissions..."
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName "nick.seredynski@mobysoft.com"

# Find shared mailboxes where the source user has permissions
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
                -User $targetUserUPN `
                -AccessRights $permission.AccessRights `
                -InheritanceType $permission.InheritanceType
            Write-Host "Copied permissions for shared mailbox $($sharedMailbox.PrimarySmtpAddress)."
        } catch {
            Write-Warning "Could not copy permissions for shared mailbox $($sharedMailbox.PrimarySmtpAddress): $_"
        }
    }
}

# Step 4: Copy Custom Attributes (Including Company Name and Manager)
Write-Host "Copying custom attributes (e.g., Department, Job Title, Company Name, Manager)..."
try {
    # Copy standard attributes
    Set-AzureADUser -ObjectId $targetUser.ObjectId `
        -Department $sourceUser.Department `
        -JobTitle $sourceUser.JobTitle `
        -City $sourceUser.City `
        -Country $sourceUser.Country `
        -CompanyName $sourceUser.CompanyName
    Write-Host "Copied standard attributes successfully."
} catch {
    Write-Warning "Could not copy standard attributes: $_"
}

# Copy Manager
Write-Host "Copying manager information..."
try {
    $manager = Get-AzureADUserManager -ObjectId $sourceUser.ObjectId
    if ($manager) {
        Set-AzureADUserManager -ObjectId $targetUser.ObjectId -RefObjectId $manager.ObjectId
        Write-Host "Copied manager: $($manager.DisplayName)"
    } else {
        Write-Host "Source user has no manager assigned."
    }
} catch {
    Write-Warning "Could not copy manager: $_"
}

# Step 5: Copy Distribution List Memberships
Write-Host "Copying distribution list memberships..."
$distributionLists = Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup | Where-Object {
    (Get-DistributionGroupMember -Identity $_.PrimarySmtpAddress | Where-Object { $_.PrimarySmtpAddress -eq $sourceUserUPN })
}

foreach ($dl in $distributionLists) {
    try {
        Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $targetUserUPN
        Write-Host "Added $targetUserUPN to distribution list $($dl.DisplayName)."
    } catch {
        if ($_.Exception.Message -like "*ResourceNotFound*") {
            Write-Warning "Distribution list $($dl.DisplayName) does not exist or cannot be accessed. Skipping."
        } else {
            Write-Warning "Could not add $targetUserUPN to distribution list $($dl.DisplayName): $_"
        }
    }
}

Write-Host "Permissions copying process completed successfully (excluding roles)."
