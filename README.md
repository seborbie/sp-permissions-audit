# SharePoint Online Permissions Audit Script

It is well known that SharePoint permissions are notoriously difficult to manage. This script is designed to help you audit permissions across your SharePoint Online sites.

This script is based on a fork of [sp-permissions-audit](https://github.com/scottmckendry/sp-permissions-audit). The original version relied on the PnP PowerShell module, which introduced significant overhead and did not support concurrent operations. This version has been fully refactored to use the Microsoft Graph API and SharePoint REST API, enabling robust concurrency and greatly reducing execution time. Depending on your system and the number of available threads or cores, the script typically runs 20x–40x faster than the original, while producing identical output. You are welcome to use and modify this script as needed. Please note that it is provided as-is, without warranty, see the license file for details.

## ✨ Features

-   Audit permissions for all sites in a SharePoint Online tenant - all the way down to list and library level.
-   Capture permissions granted to Security (Entra ID) and Microsoft 365 groups.
-   Uses a modern authentication flow that does not require a user to be logged in or have access to all sites in the tenant.
 -   Tracks per-user runtime and appends the total seconds to the CSV as a summary row.
 -   Optional transcript logging to a file via `-Log` (with `-AppendLog` support).
-   Enumerate all users in a tenant via Microsoft Graph `Directory.Read.All` using `Get-TenantUsers.ps1`.

## ⚡ Performance

The script supports parallel processing in PowerShell 7 using direct REST API calls. Control the level of concurrency with `-ThrottleLimit`. Choose a value that suits your environment and available capacity.

## 📝 Output

The script will output a CSV file with the following columns:

| Column Name       | Description                                                                                       |
| ----------------- | ------------------------------------------------------------------------------------------------- |
| UserPrincipalName | The user's UPN/email address                                                                      |
| SiteUrl           | The URL of the site                                                                               |
| SiteAdmin         | Is the user a site admin?                                                                         |
| GroupName         | If the user is not a site admin, what SharePoint group are they in? (also captures sharing links) |
| PermissionLevel   | The permission level granted to the SharePoint group, e.g full control, read, edit etc.           |
| ListName          | The title of a list or library where the user has unique permissions.                             |
| ListPermission    | The permission level granted to the user on the list or library.                                  |
| TotalRuntimeSeconds | The total runtime (in seconds) for enumerating a user's permissions. Populated only in the final summary row for that user; null for other rows. |

In addition to the detailed rows, the script appends a final summary row per user with `TotalRuntimeSeconds` populated and other fields left null.

## 🚀 Getting Started

### Prerequisites

-   Global Adminstrator Role
-   PowerShell 7 or later with the latest version of [MSAL.PS](https://github.com/AzureAD/MSAL.PS/) installed.
-   A self-signed certificate for use with the app registration. See [this article](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread) for more information.

```powershell
Install-Module -Name MSAL.PS -Scope CurrentUser
```

### Create an Entra ID App Registration

Follow the steps in [this article](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread) to create an app registration in Azure AD. Make sure you grant the app the following permissions.

**Graph API**

-   Sites.Read.All
-   Directory.Read.All

**SharePoint API**

-   Sites.FullControl.All
-   User.Read.All

## Usage

The intention is for this script to be called by a parent script that will pass in the required parameters. This allows you to run the script against multiple users and potentially multiple tenants.
Below is an example of how you might call the script. This example assumes you've exported a list of users from Entra ID into a CSV file that would then get imported for the purpose of this script.

```powershell
# audit.ps1 - Create in the same directory as Get-SharePointTenantPermissions.ps1

$tenantName = "contoso" # The name of your tenant, e.g. contoso.sharepoint.com
$csvPath = "C:\temp\permissions.csv" # The path to the output CSV file
$clientID = "00000000-0000-0000-0000-000000000000" # The client ID of the app registration
$certificatePath = "C:\temp\certificate.pfx" # The path to the certificate file
$append = $true # Should the script append to the CSV file or overwrite it?
$logPath = "C:\logs\sp-permissions\audit.log" # Optional: transcript log path
$certPassword = Read-Host -AsSecureString -Prompt "Enter PFX password" # Optional: only if your PFX is password-protected
$throttleLimit = 1 # Number of parallel threads (set higher to process sites in parallel)

$users = Import-Csv -Path "C:\temp\users.csv"

foreach ($user in $users) {
    .\Get-SharePointTenantPermissions.ps1 `
        -TenantName $tenantName `
        -CsvPath $csvPath `
        -ClientID $clientID `
        -CertificatePath $certificatePath `
        -CertificatePassword $certPassword ` # Optional: include if your PFX has a password
        -Append:$append `
        -UserEmail $user.UserPrincipalName `
        -ThrottleLimit $throttleLimit `    # Optional: adjust parallel processing (default: 10)
        -Log $logPath `                    # Optional: write output to a transcript log
        -AppendLog                         # Optional: append to existing log instead of overwriting
}
```

### Parallel processing

-   **-ThrottleLimit <int>**: Controls the number of parallel threads used for processing sites (default: 1 - sequential). Increase to process sites faster using multiple threads.

### Optional logging

-   **-Log <path>**: If supplied, the script starts a transcript and writes all output to the specified file (creating the directory if needed). If not supplied, output goes to the console as usual.
-   **-AppendLog**: When used with `-Log`, appends to the existing log file. If omitted, any existing log file at the given path is overwritten.

When `-Log` is supplied, console output is minimized to major events to reduce console I/O and improve performance when processing many users:

-   **Major console events**: start/end per user, writing results to CSV, total runtime seconds.
-   **Detailed events** (e.g., per-site progress like `Processing https://contoso.sharepoint.com/sites/XYZ (10 of 372)`, and permission findings) are written to the log file only.

### Password-protected PFX certificates

-   If your certificate PFX is protected with a password, pass it via the new `-CertificatePassword` parameter as a `SecureString`.
-   Interactive example: `Read-Host -AsSecureString -Prompt "Enter PFX password"` and pass the variable to `-CertificatePassword`.
-   Automated example (less secure): `ConvertTo-SecureString 'PlainTextPassword' -AsPlainText -Force`.
-   If your PFX has no password, omit `-CertificatePassword`.

## 🤝 Contributing

Contributions, issues and feature requests are welcome!

TODO:

-   [ ] Replace [MSAL.PS](https://github.com/AzureAD/MSAL.PS) cmdlets with a non-deprecated alternative
