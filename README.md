SYNOPSIS

    Exports SharePoint site file permissions to CSV

DESCRIPTION

    This script connects to Microsoft Graph and exports all file permissions from a SharePoint site to a CSV file.
    It processes all document libraries and includes detailed permission information.

PARAMETER SiteUrl

    The URL of the SharePoint site to analyze

PARAMETER OutputCsv

    The output CSV file path (default: SharedFilesPermissions.csv)

PARAMETER VerboseOutput

    Enable verbose output for detailed processing information

PARAMETER IncludeSystemFiles

    Include system files in the analysis (default: excluded)

PARAMETER MaxDepth

    Maximum folder depth to recurse (default: 10)

EXAMPLE

    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

EXAMPLE

    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -OutputCsv "permissions.csv" -VerboseOutput

EXAMPLE

    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -ClientSecret "your-client-secret"

EXAMPLE

    Using existing global token:
    $global:graphAPIToken = "your-access-token-here"
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

NOTES

    Requires Microsoft.Graph PowerShell module
    Requires Sites.Read.All, Files.Read.All, User.Read.All permissions
    
     Authentication Methods:
    1. App registration (ClientId, TenantId, ClientSecret) - Non-interactive, optionally uses existing global token
    2. Interactive authentication - Prompts for user login
    
    Global Token Usage:
    - If $global:graphAPIToken is set, it will be used with app registration authentication
    - This allows reusing existing tokens without separate authentication flows
    
    Troubleshooting App Registration Issues:
    - Ensure APPLICATION permissions are granted (not delegated)
    - Admin consent must be granted for all permissions
    - Wait 5-10 minutes after granting consent for permissions to propagate
    - Verify the app registration is in the correct tenant
    - Check that the client secret hasn't expired
