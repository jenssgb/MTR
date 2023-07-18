# Import required modules if not already imported
$requiredModules = @('MSOnline', 'AzureAD', 'ExchangeOnlineManagement')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -Force
    }
}

# Import modules
Import-Module $requiredModules

# Connect to Exchange
$UserCredential = Get-Credential
Connect-ExchangeOnline -Credential $UserCredential

# List all available rooms
Get-Mailbox -RecipientTypeDetails RoomMailbox | Format-Table Name, Alias, Database, ProhibitSendQuota, ExternalDirectoryObjectId

# Extract the domain
$Domain = $UserCredential.UserName.Split("@")[1]

# Prompt for action to take
$Action = Read-Host -Prompt 'Enter the action to take (1 for Create, 2 for Delete)'

# Create or delete mailbox
if ($Action -eq '1') {
    $Alias = Read-Host -Prompt 'Enter the Alias'
    $Password = Read-Host -Prompt 'Enter the Room Mailbox Password' -AsSecureString
    $mailboxParams = @{
        MicrosoftOnlineServicesID = "$Alias@$Domain"
        Name = $Alias
        Alias = $Alias
        Room = $true
        EnableRoomMailboxAccount = $true
        RoomMailboxPassword = $Password
    }
    New-Mailbox @mailboxParams
    Start-Sleep -Seconds 30

    # Wait and retry if mailbox is not immediately available
    for ($i = 0; $i -lt 20; $i++) {
        try {
            $calendarParams = @{
                Identity = $Alias
                AutomateProcessing = 'AutoAccept'
                AddOrganizerToSubject = $true
                DeleteComments = $false
                DeleteSubject = $false
                ProcessExternalMeetingMessages = $true
                RemovePrivateProperty = $false
                AddAdditionalResponse = $true
                AdditionalResponse = 'This is a Microsoft Teams Meeting room!'
            }
            Set-CalendarProcessing @calendarParams
            break
        }
        catch {
            if ($i -eq 19) {
                throw
            }
            else {
                Start-Sleep -Seconds 60
            }
        }
    }

    # Connect to Azure AD
    Connect-AzureAD -Credential $UserCredential
	
	# Connect to MSOnline service
	Connect-MsolService -Credential $UserCredential
	

    # Never Expire Password
    Set-AzureADUser -ObjectId "$Alias@$Domain" -PasswordPolicies DisablePasswordExpiration

    # Assign license
    $AccountSkuId = 'M365x45115115:Microsoft_Teams_Rooms_Pro'
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuId
    try {
        Set-MsolUserLicense -UserPrincipalName "$Alias@$Domain" -AddLicenses $AccountSkuId -LicenseOptions $LicenseOptions
    }
    catch {
        Write-Host "No available licenses for $AccountSkuId. Continuing with the rest of the operations."
    }

    # Output the new mailbox info
    Write-Host "Created mailbox: $Alias@$Domain"
    Write-Host "Password: $(ConvertFrom-SecureString -SecureString $Password)"
}


# Lösche vorhandene Mailbox
elseif ($Action -eq '2') {
    $Alias = Read-Host -Prompt 'Enter the Alias'
    Remove-Mailbox -Identity "$Alias@$Domain" -Confirm:$false

    # Verbinde mit Azure AD
    Connect-AzureAD

    # Gib Lizenz frei
    $AccountSkuId = "M365x45115115:Microsoft_Teams_Rooms_Pro"
    try {
        Set-MsolUserLicense -UserPrincipalName "$Alias@$Domain" -RemoveLicenses $AccountSkuId
    }
    catch {
        Write-Host "Failed to release the license for $Alias@$Domain"
    }
}




else {
    Write-Host "Invalid action. Please enter 1 for Create or 2 for Delete."
}

# Output the result
Write-Host "End of script."

