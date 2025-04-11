[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Mailbox,
    
    [Parameter(Mandatory = $true)]
    [string]$User,
    
    [Parameter(Mandatory = $false)]
    [switch]$FullAccess = $false
)

# Check if Exchange Online module is installed and connected
function Ensure-ExchangeOnlineConnection {
    # Check if the ExchangeOnlineManagement module is installed
    if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Write-Host "Exchange Online Management module is not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
            Write-Host "Exchange Online Management module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install Exchange Online Management module. Error: $_"
            exit 1
        }
    }

    # Check if already connected to Exchange Online
    try {
        $connectionStatus = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
        Write-Host "Already connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        # Not connected, attempt to connect
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        try {
            Connect-ExchangeOnline -ErrorAction Stop
            Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to connect to Exchange Online. Error: $_"
            exit 1
        }
    }
}

# Main script execution
try {
    # Ensure we're connected to Exchange Online
    Ensure-ExchangeOnlineConnection

    # Verify the mailbox exists
    try {
        $mailboxCheck = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        Write-Host "Target mailbox '$Mailbox' found." -ForegroundColor Green
    }
    catch {
        Write-Error "Target mailbox '$Mailbox' not found. Please verify the mailbox name or email address."
        exit 1
    }

    # Verify the user exists
    try {
        $userCheck = Get-User -Identity $User -ErrorAction Stop
        Write-Host "User '$User' found." -ForegroundColor Green
    }
    catch {
        Write-Error "User '$User' not found. Please verify the user name or email address."
        exit 1
    }
    
    # Determine the appropriate access rights based on parameter
    if ($FullAccess) {
        $accessRights = "FullAccess"
        Write-Host "Using FullAccess permission level for OWA compatibility." -ForegroundColor Yellow
    } else {
        $accessRights = "ReadPermission"
        Write-Host "Using ReadPermission permission level. Note: If OWA access fails, try running the script with -FullAccess." -ForegroundColor Yellow
    }

    # Add user to mailbox with appropriate permissions and disable automapping
    Write-Host "Adding user '$User' to mailbox '$Mailbox' with $accessRights permissions and disabling automapping..." -ForegroundColor Yellow
    
    try {
        # First, check if there are existing permissions and remove them to avoid conflicts
        $existingPermissions = Get-MailboxPermission -Identity $Mailbox -User $User -ErrorAction SilentlyContinue
        if ($existingPermissions) {
            Write-Host "Found existing permissions. Removing to avoid conflicts..." -ForegroundColor Yellow
            Remove-MailboxPermission -Identity $Mailbox -User $User -AccessRights $existingPermissions.AccessRights -Confirm:$false -ErrorAction Stop
            Write-Host "Existing permissions removed." -ForegroundColor Green
        }
        
        # Now add the appropriate permission
        $addResult = Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights $accessRights -AutoMapping $false -ErrorAction Stop
        Write-Host "User added with $accessRights permissions and automapping disabled." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to add mailbox permission. Error: $_"
        # Try a different approach if there's an error
        try {
            Write-Host "Attempting alternative permission setting method..." -ForegroundColor Yellow
            if ($FullAccess) {
                # Try adding with explicit FullAccess for mailbox
                Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights FullAccess -InheritanceType All -AutoMapping $false
            }
            else {
                # If ReadPermission failed, try adding specific folder permissions for OWA compatibility
                Add-MailboxFolderPermission -Identity "$($Mailbox):\Calendar" -User $User -AccessRights Reviewer
                Add-MailboxFolderPermission -Identity "$($Mailbox):\Inbox" -User $User -AccessRights Reviewer
                Add-MailboxFolderPermission -Identity "$($Mailbox):\Sent Items" -User $User -AccessRights Reviewer
                Add-MailboxFolderPermission -Identity "$($Mailbox):\Drafts" -User $User -AccessRights Reviewer
            }
            Write-Host "Permissions added using alternative method." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to set permissions using alternative method. Error: $_"
        }
    }

    # Enable Outlook on the Web access for the user
    Write-Host "Enabling Outlook on the Web access for user..." -ForegroundColor Yellow
    
    try {
        Set-CASMailbox -Identity $User -OWAEnabled $true -ErrorAction Stop
        Write-Host "Outlook on the Web access enabled for user '$User'." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to enable Outlook on the Web access. Error: $_"
    }

    # Add specific folder permissions to ensure OWA access works
    Write-Host "Setting folder-level permissions for OWA compatibility..." -ForegroundColor Yellow
    try {
        # The DefaultFolderPermission for OWA is often "Reviewer" which allows read access
        Add-MailboxFolderPermission -Identity "$($Mailbox):\\" -User $User -AccessRights Reviewer -ErrorAction SilentlyContinue
        Write-Host "Root folder permission set." -ForegroundColor Green
    }
    catch {
        Write-Host "Note: Could not set root folder permission. This is sometimes normal. Error: $_" -ForegroundColor Yellow
    }

    # Verify all changes were applied successfully
    Write-Host "`nVerifying changes..." -ForegroundColor Yellow
    
    # Check mailbox permissions
    $verifyPermissions = Get-MailboxPermission -Identity $Mailbox -User $User
    if ($verifyPermissions) {
        Write-Host "✓ User '$User' has been granted the following permissions to mailbox '$Mailbox':" -ForegroundColor Green
        $verifyPermissions | Format-Table User, AccessRights, IsInherited, Deny -AutoSize
        
        # Check specifically for the requested access rights
        if ($verifyPermissions.AccessRights -contains $accessRights) {
            Write-Host "✓ $accessRights permission is confirmed." -ForegroundColor Green
        }
        else {
            Write-Host "❌ $accessRights was not found in the granted permissions." -ForegroundColor Red
        }
    }
    else {
        Write-Host "❌ Failed to verify mailbox permissions." -ForegroundColor Red
        
        # Check folder permissions as an alternative
        Write-Host "Checking folder permissions instead..." -ForegroundColor Yellow
        $folderPermissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\Inbox" -User $User -ErrorAction SilentlyContinue
        if ($folderPermissions) {
            Write-Host "✓ Found folder-level permissions:" -ForegroundColor Green  
            $folderPermissions | Format-Table User, AccessRights -AutoSize
        }
    }
    
    # Check OWA access
    try {
        $verifyCAS = Get-CASMailbox -Identity $User -ErrorAction Stop
        if ($verifyCAS -and $verifyCAS.OWAEnabled) {
            Write-Host "✓ Outlook on the Web access is enabled for user '$User'" -ForegroundColor Green
        }
        else {
            Write-Host "❌ Outlook on the Web access appears to be disabled for user '$User'" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "❌ Failed to verify Outlook on the Web access." -ForegroundColor Red
    }

    Write-Host "`nOperation completed." -ForegroundColor Green
    Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Yellow
    Write-Host "1. It may take 15-60 minutes for permissions to fully propagate in Exchange Online." -ForegroundColor Yellow
    Write-Host "2. If OWA access still doesn't work after waiting, try running this script again with -FullAccess switch." -ForegroundColor Yellow
    Write-Host "3. The user may need to sign out of OWA completely and sign back in for changes to take effect." -ForegroundColor Yellow
}
catch {
    Write-Error "An error occurred during script execution: $_"
    exit 1
}