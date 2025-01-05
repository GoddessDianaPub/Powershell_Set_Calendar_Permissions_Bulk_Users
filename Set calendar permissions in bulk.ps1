# Define Manager UPN and Access Rights
$sourceUPN = "user@domain.com"
$accessRights = @("accessRights")

# Define the users as an array
$DestinationUPNs = @(
    "user1@domain.com",
    "user2@domain.com",
    "user3@domain.com"
)

# Loop through each secretary and configure permissions
foreach ($DestinationUPN in $DestinationUPNs) {
    $calendarIdentity = "${sourceUPN}:\Calendar"

    # Add permissions
    Write-Host "Granting $accessRights access to $DestinationUPNs for $calendarIdentity..."
    try {
        Set-MailboxFolderPermission -Identity $calendarIdentity -User $DestinationUPN -AccessRights $accessRights
        Write-Host "Successfully granted $accessRights to $DestinationUPNs." -ForegroundColor Green
    } catch {
        Write-Host "Failed to grant permissions to $DestinationUPNs. Error: $_" -ForegroundColor Red
    }
}

Get-MailboxFolderPermission -Identity $calendarIdentity | Format-Table FolderName, User, AccessRights -AutoSize