# Exchange Mailbox Access Script

A PowerShell script that manages mailbox access permissions in Exchange Online with a focus on configuring proper access for Outlook on the Web (OWA).

## Features

- Adds users to mailboxes with configurable permission levels
- Disables Auto-Mapping to prevent the shared mailbox from auto-loading in Outlook
- Enables Outlook on the Web (OWA) access for added users
- Handles folder-level permissions for complete OWA compatibility
- Provides verification of all changes applied
- Contains built-in error handling and troubleshooting

## Requirements

- PowerShell 5.1 or higher
- Exchange Online PowerShell V2 module
- Appropriate Exchange Online administrative permissions

## Installation

1. Download the `Add-MailboxAccess.ps1` script to your local system
2. If you haven't installed the Exchange Online PowerShell module, the script will attempt to install it automatically

## Usage

### Basic Usage - Read Permission

```powershell
.\Add-MailboxAccess.ps1 -Mailbox "shared.mailbox@company.com" -User "john.doe@company.com"
```

This grants the user read-only access to the mailbox.

### Full Access Permission

```powershell
.\Add-MailboxAccess.ps1 -Mailbox "shared.mailbox@company.com" -User "john.doe@company.com" -FullAccess
```

This grants the user full access to the mailbox, which is sometimes required for complete OWA functionality.

### Verbose Mode

```powershell
.\Add-MailboxAccess.ps1 -Mailbox "shared.mailbox@company.com" -User "john.doe@company.com" -Verbose
```

Provides detailed output about each step of the process, useful for troubleshooting. Note that `-Verbose` is a built-in PowerShell common parameter that works with any cmdlet or script that uses `[CmdletBinding()]`.

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-Mailbox` | String | Yes | The target mailbox to grant access to. Can be an email address, UPN, or display name. |
| `-User` | String | Yes | The user who should receive access. Can be an email address, UPN, or display name. |
| `-FullAccess` | Switch | No | When specified, grants full access instead of read-only access. |
| `-Verbose` | Switch | No | Built-in PowerShell parameter that displays detailed diagnostic information during execution. |

## Common Issues and Solutions

### OWA Access Issues

If the user receives "You might not have permission to perform this action" in OWA:

1. **Wait for permission propagation**: Exchange Online permissions typically take 15-60 minutes to fully propagate
2. **Try with FullAccess**: Run the script again with the `-FullAccess` parameter
3. **Sign out and back in**: Have the user completely sign out of OWA and sign back in
4. **Check administrator permissions**: Ensure the account running the script has appropriate Exchange admin rights

### Permission Conflicts

If you receive errors about existing permissions:

1. **Remove existing permissions first**: The script will attempt to do this automatically
2. **Check for conflicting permissions**: Use `Get-MailboxPermission -Identity "mailbox" -User "user"` to check current permissions
3. **Run with -Verbose**: Get detailed diagnostic information about what might be causing conflicts

## License

This script is provided under the MIT License. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request
