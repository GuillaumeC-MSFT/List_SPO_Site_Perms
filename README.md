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

NOTES

    Requires Microsoft.Graph PowerShell module
    Requires Sites.Read.All, Files.Read.All, User.Read.All permissions
