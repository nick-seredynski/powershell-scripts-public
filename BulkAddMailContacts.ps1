# Connect to Exchange Online
Connect-ExchangeOnline

# Path to CSV, names must match columns in CSV, referenced in script as $c.FirstName:
#	FirstName	|	LastName	|	Email
$contacts = Import-Csv ""C:\User\Scripts\Contacts.csv""

# for loop for looping through CSV file $c is the name of each iteration
foreach ($c in $contacts) {

    # Generate DisplayName based on each iteration
    $displayName = "$($c.FirstName) $($c.LastName)"

    # Create a valid alias:
    $alias = ($c.FirstName + "." + $c.LastName).ToLower()                        # firstname.lastname
    $alias = $alias -replace '[^a-z0-9._-]', ''  # strip anything invalid        # remove invalid characters
    $alias = $alias.Trim(".")                    # no leading or trailing dot

    # Safety check: ensure alias is not empty
    if ($alias -eq "") {
        Write-Warning "Alias could not be generated for $displayName. Skipping."
        continue
    }

    # Check for existing external email
    $existingEmail = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$($c.Email)'" -ErrorAction SilentlyContinue
    if ($existingEmail) {
        Write-Warning "External email '$($c.Email)' already exists. Skipping $displayName."
        continue
    }

    try {
        New-MailContact `
            -Name $displayName `
            -DisplayName $displayName `
            -FirstName $c.FirstName `
            -LastName $c.LastName `
            -Alias $alias `
            -ExternalEmailAddress $c.Email `
            -ErrorAction Stop

        Write-Host "Created contact: $displayName" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create contact $displayName. Error: $_" -ForegroundColor Red
    }
}

# Disconnect from ExchangeOnline
Disconnect-ExchangeOnline -Confirm:$false

# Prompt user to press any key to exit - this allows admin to read errors and success logs
Read-Host "Press Enter to exit"
