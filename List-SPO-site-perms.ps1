<#
.SYNOPSIS
    Exports SharePoint site file permissions to CSV

.DESCRIPTION
    This script connects to Microsoft Graph and exports all file permissions from a SharePoint site to a CSV file.
    It processes all document libraries and includes detailed permission information.

.PARAMETER SiteUrl
    The URL of the SharePoint site to analyze

.PARAMETER OutputCsv
    The output CSV file path (default: SharedFilesPermissions.csv)

.PARAMETER VerboseOutput
    Enable verbose output for detailed processing information

.PARAMETER IncludeSystemFiles
    Include system files in the analysis (default: excluded)

.PARAMETER MaxDepth
    Maximum folder depth to recurse (default: 10)

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -OutputCsv "permissions.csv" -VerboseOutput

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -ClientSecret "your-client-secret"

.EXAMPLE
    # Using existing global token
    $global:graphAPIToken = "your-access-token-here"
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

.NOTES
    Requires Microsoft.Graph PowerShell module
    Requires Sites.Read.All, Files.Read.All, User.Read.All permissions
    
    Authentication Methods (in order of preference):
    1. Global token ($global:graphAPIToken) - Uses existing access token
    2. App registration (ClientId, TenantId, ClientSecret) - Non-interactive
    3. Interactive authentication - Prompts for user login
    
    Troubleshooting App Registration Issues:
    - Ensure APPLICATION permissions are granted (not delegated)
    - Admin consent must be granted for all permissions
    - Wait 5-10 minutes after granting consent for permissions to propagate
    - Verify the app registration is in the correct tenant
    - Check that the client secret hasn't expired
#>

param (
    [Parameter(Mandatory, HelpMessage="Enter the SharePoint site URL")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,
    
    [Parameter(HelpMessage="Output CSV file path")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputCsv = "SharedFilesPermissions.csv",
    
    [Parameter(HelpMessage="Enable verbose output")]
    [switch]$VerboseOutput,
    
    [Parameter(HelpMessage="Include system files in analysis")]
    [switch]$IncludeSystemFiles,
    
    [Parameter(HelpMessage="Maximum folder depth to recurse")]
    [ValidateRange(1, 50)]
    [int]$MaxDepth = 10,
    
    [Parameter(HelpMessage="Application (Client) ID for app registration authentication")]
    [string]$ClientId,
    
    [Parameter(HelpMessage="Tenant ID for app registration authentication")]
    [string]$TenantId,
    
    [Parameter(HelpMessage="Client secret for app registration authentication")]
    [string]$ClientSecret
)

# Initialize variables
$startTime = Get-Date
$processedFiles = 0
$totalPermissions = 0
$systemFilePatterns = @("~\$", "\.tmp$", "Forms/", "_vti_", "SitePages/", "SiteAssets/", "Style Library/")

# Ensure output directory exists
$outputDir = Split-Path $OutputCsv -Parent
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Function to check if file should be included
function Should-IncludeFile {
    param($fileName, $filePath)
    
    if ($IncludeSystemFiles) {
        return $true
    }
    
    foreach ($pattern in $systemFilePatterns) {
        if ($fileName -match $pattern -or $filePath -match $pattern) {
            return $false
        }
    }
    return $true
}

# Function to get files recursively
function Get-FilesRecursively {
    param($DriveId, $ParentId = $null, $CurrentDepth = 0, $MaxDepth = 10)
    
    if ($CurrentDepth -ge $MaxDepth) {
        return @()
    }
    
    $allFiles = @()
    
    try {
        if ($ParentId) {
            $items = Get-MgDriveItem -DriveId $DriveId -DriveItemId $ParentId -ExpandProperty "children" -ErrorAction Stop
            $children = $items.Children
        } else {
            $children = Get-MgDriveItem -DriveId $DriveId -Filter "file ne null or folder ne null" -ErrorAction Stop
        }
        
        foreach ($item in $children) {
            if ($item.File) {
                # It's a file
                if (Should-IncludeFile -fileName $item.Name -filePath $item.WebUrl) {
                    $allFiles += $item
                }
            } elseif ($item.Folder -and $CurrentDepth -lt $MaxDepth) {
                # It's a folder - recurse into it
                if ($VerboseOutput) {
                    Write-Output "  Scanning folder: $($item.Name) (depth: $($CurrentDepth + 1))"
                }
                $allFiles += Get-FilesRecursively -DriveId $DriveId -ParentId $item.Id -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            }
        }
    } catch {
        Write-Output "Failed to retrieve items at depth $CurrentDepth`: $_"
    }
    
    return $allFiles
}

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Output "Installing Microsoft.Graph module..."
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
try {
    # Check authentication methods in order of preference
    if ($global:graphAPIToken) {
        # Method 1: Use existing global token
        Write-Output "Connecting to Microsoft Graph using existing token..."
        $secureToken = ConvertTo-SecureString $global:graphAPIToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
        Write-Output "Connected to Microsoft Graph using existing token."
    } elseif ($ClientId -and $TenantId -and $ClientSecret) {
        # Method 2: App registration with client secret
        Write-Output "Connecting to Microsoft Graph using app registration..."
        Write-Output "Client ID: $ClientId"
        Write-Output "Tenant ID: $TenantId"
        Write-Output "Client Secret: $($ClientSecret.Substring(0, 5))..." # Show first 5 chars only
        
        # Validate parameters format
        if ([string]::IsNullOrWhiteSpace($ClientId) -or $ClientId.Length -ne 36) {
            Write-Output "ERROR: ClientId appears invalid. Should be a 36-character GUID."
            Write-Output "Format: 12345678-1234-1234-1234-123456789012"
            exit 1
        }
        
        if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId.Length -ne 36) {
            Write-Output "ERROR: TenantId appears invalid. Should be a 36-character GUID."
            Write-Output "Format: 87654321-4321-4321-4321-210987654321"
            exit 1
        }
        
        if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
            Write-Output "ERROR: ClientSecret is empty."
            exit 1
        }
        
        # Convert client secret to secure string
        try {
            $secureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $clientCredential = New-Object System.Management.Automation.PSCredential($ClientId, $secureClientSecret)
        } catch {
            Write-Output "ERROR: Failed to create credentials: $_"
            exit 1
        }
        
        # Disconnect any existing sessions
        try { 
            Disconnect-MgGraph -ErrorAction SilentlyContinue 
        } catch { 
            # Ignore disconnect errors
        }
        
        # Connect using client credentials with detailed error handling
        try {
            Write-Output "Attempting connection to Microsoft Graph..."
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientCredential -NoWelcome -ErrorAction Stop
            Write-Output "✓ Connect-MgGraph command executed successfully"
        } catch {
            Write-Output "✗ Connect-MgGraph failed:"
            Write-Output "Error: $($_.Exception.Message)"
            Write-Output "Category: $($_.CategoryInfo.Category)"
            Write-Output "FullyQualifiedErrorId: $($_.FullyQualifiedErrorId)"
            
            # Common error patterns and solutions
            if ($_.Exception.Message -match "AADSTS7000215") {
                Write-Output "`nDIAGNOSIS: Invalid client secret provided"
                Write-Output "SOLUTION: Regenerate client secret in Azure Portal"
            } elseif ($_.Exception.Message -match "AADSTS7000222") {
                Write-Output "`nDIAGNOSIS: Client secret has expired"
                Write-Output "SOLUTION: Create new client secret in Azure Portal"
            } elseif ($_.Exception.Message -match "AADSTS700016") {
                Write-Output "`nDIAGNOSIS: Invalid application ID (ClientId)"
                Write-Output "SOLUTION: Verify ClientId from app registration overview"
            } elseif ($_.Exception.Message -match "AADSTS90002") {
                Write-Output "`nDIAGNOSIS: Invalid tenant ID"
                Write-Output "SOLUTION: Verify TenantId from app registration overview"
            } else {
                Write-Output "`nCOMMON ISSUES TO CHECK:"
                Write-Output "1. App registration exists in the correct tenant"
                Write-Output "2. Client secret is the VALUE not the Secret ID"
                Write-Output "3. Client secret hasn't expired"
                Write-Output "4. ClientId and TenantId are correct GUIDs"
                Write-Output "5. App registration has required permissions"
            }
            exit 1
        }
        
        # Verify the connection worked
        Write-Output "Verifying connection context..."
        $context = Get-MgContext
        if (-not $context -or -not $context.Account) {
            Write-Output "✗ Authentication failed - no valid context found"
            Write-Output "Context details: $($context | ConvertTo-Json -Depth 2)"
            Write-Output "`nThis usually indicates:"
            Write-Output "1. The app registration lacks required APPLICATION permissions"
            Write-Output "2. Admin consent has not been granted"
            Write-Output "3. Permissions haven't propagated yet (wait 5-10 minutes)"
            exit 1
        }
        
        Write-Output "✓ Connection context verified"
        Write-Output "Account: $($context.Account)"
        Write-Output "AuthType: $($context.AuthType)"
        Write-Output "TenantId: $($context.TenantId)"
        
        # Test basic Graph API access
        Write-Output "Testing basic Graph API access..."
        try {
            $profile = Get-MgProfile -ErrorAction Stop
            Write-Output "✓ Basic Graph API access: SUCCESS"
        } catch {
            Write-Output "✗ Basic Graph API access: FAILED"
            Write-Output "Error: $($_.Exception.Message)"
            Write-Output "This usually indicates missing permissions or admin consent issues"
        }
        
        # Test Sites.Read.All permission specifically
        Write-Output "Testing Sites.Read.All permission..."
        try {
            $testSites = Get-MgSite -Top 1 -ErrorAction Stop
            Write-Output "✓ Sites.Read.All permission: SUCCESS"
            Write-Output "Found $($testSites.Count) site(s) - permission is working"
        } catch {
            Write-Output "✗ Sites.Read.All permission: FAILED"
            Write-Output "Error: $($_.Exception.Message)"
            Write-Output ""
            Write-Output "CRITICAL: This permission is required for SharePoint access"
            Write-Output "TO FIX:"
            Write-Output "1. Go to Azure Portal → App Registrations → [Your App] → API Permissions"
            Write-Output "2. Add Microsoft Graph → Application permissions → Sites.Read.All"
            Write-Output "3. Click 'Grant admin consent for [your organization]'"
            Write-Output "4. Wait 5-10 minutes for permissions to propagate"
            exit 1
        }
        
        Write-Output "✓ App registration authentication fully verified"
        
    } else {
        # Method 3: Interactive authentication (existing behavior)
        Write-Output "Connecting to Microsoft Graph interactively..."
        Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"
        Write-Output "Connected to Microsoft Graph."
    }
    
    # Final verification that we're connected
    $context = Get-MgContext
    if (-not $context -or -not $context.Account) {
        Write-Output "FATAL ERROR: Failed to establish Microsoft Graph connection"
        exit 1
    }
    
    if ($VerboseOutput) {
        Write-Output "Final Graph context details:"
        Write-Output "• Account: $($context.Account)"
        Write-Output "• Auth Type: $($context.AuthType)"
        Write-Output "• Scopes: $($context.Scopes -join ', ')"
        Write-Output "• Tenant ID: $($context.TenantId)"
        Write-Output "• Client ID: $($context.ClientId)"
        Write-Output "• Certificate Thumbprint: $($context.CertificateThumbprint)"
    }
    
} catch {
    Write-Output "FATAL ERROR: Failed to connect to Microsoft Graph: $_"
    Write-Output "Error details: $($_.Exception.Message)"
    exit 1
}

# Validate Site URL
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Output "No SharePoint site URL provided. Exiting."
    exit 1
}

try {
    $uri = [System.Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath.TrimStart("/")
} catch {
    Write-Output "Invalid site URL format. Exiting."
    exit 1
}

$hostname = $uri.Hostname

# Get Site Object
try {
    # Try to get the site directly by URL
    $site = Get-MgSite -Search "${hostname}:${sitePath}" | Where-Object { $_.WebUrl -eq $SiteUrl }

    if (-not $site) {
        # Fallback: try searching by hostname only
        $site = Get-MgSite -Search $hostname | Where-Object { $_.WebUrl -eq $SiteUrl }
    }

    if (-not $site) {
        Write-Output "Site not found. Please check the URL and try again."
        exit 1
    }
} catch {
    if ($_.Exception.Message -match "401|Unauthorized") {
        Write-Output "Access denied (401 Unauthorized)"
        Write-Output "This typically means the authentication succeeded but your app registration lacks the required permissions."
        Write-Output ""
        Write-Output "To fix this issue:"
        Write-Output "1. Go to Azure Portal → App Registrations → [Your App] → API Permissions"
        Write-Output "2. Ensure these APPLICATION permissions are added:"
        Write-Output "   • Sites.Read.All"
        Write-Output "   • Files.Read.All"
        Write-Output "   • User.Read.All"
        Write-Output "3. Click 'Grant admin consent for [your organization]'"
        Write-Output "4. Wait a few minutes for permissions to propagate"
        Write-Output ""
        Write-Output "Current authentication context:"
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Output "   • Account: $($context.Account)"
            Write-Output "   • Auth Type: $($context.AuthType)"
            Write-Output "   • Scopes: $($context.Scopes -join ', ')"
            Write-Output "   • Tenant ID: $($context.TenantId)"
        }
        Write-Output ""
        Write-Output "Note: Application permissions require admin consent and may take time to propagate."
    } else {
        Write-Output "Failed to retrieve site: $_"
        Write-Output "Error details: $($_.Exception.Message)"
    }
    exit 1
}
$siteId = $site.Id
Write-Output "Found site: $($site.DisplayName) (ID: $siteId)"

# Get Document Libraries (Drives)
$drives = Get-MgSiteDrive -SiteId $siteId
if ($null -eq $drives -or $drives.Count -eq 0) {
    Write-Output "No document libraries found in this site."
    exit 1
}
if ($VerboseOutput) {
    Write-Output "Document libraries found:"
    $drives | ForEach-Object { Write-Output "Drive: $($_.Name) (ID: $($_.Id))" }
}

# Use ArrayList for better performance
$results = [System.Collections.ArrayList]::new()

foreach ($drive in $drives) {
    if ($VerboseOutput) { Write-Output "Processing drive: $($drive.Name)" }
    
    try {
        # Get all files recursively
        $allFiles = Get-FilesRecursively -DriveId $drive.Id -MaxDepth $MaxDepth
        
        if ($allFiles.Count -eq 0) {
            Write-Output "No files found in drive: $($drive.Name)"
            continue
        }
        
        $fileCount = 0
        foreach ($file in $allFiles) {
            $fileCount++
            $processedFiles++
            
            if ($VerboseOutput) { 
                Write-Output "Checking file: $($file.Name) ($fileCount/$($allFiles.Count))"
            } else {
                # Show progress for non-verbose mode
                if ($fileCount % 10 -eq 0) {
                    Write-Output "Processed $fileCount/$($allFiles.Count) files in $($drive.Name)..."
                }
            }
            
            try {
                $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $file.Id -ErrorAction Stop
                
                if ($permissions.Count -eq 0) {
                    # Add entry for files with no explicit permissions
                    $null = $results.Add([PSCustomObject]@{
                        DriveName      = $drive.Name
                        FileName       = $file.Name
                        FilePath       = $file.WebUrl
                        FileSize       = $file.Size
                        CreatedDateTime = $file.CreatedDateTime
                        SharedType     = "None"
                        SharedWith     = "No explicit permissions"
                        SharedEmail    = ""
                        SharedGroup    = ""
                        Identities     = ""
                        Roles          = ""
                        LastModified   = $file.LastModifiedDateTime
                        PermissionId   = ""
                        ScanDateTime   = $startTime
                    })
                } else {
                    foreach ($perm in $permissions) {
                        $totalPermissions++
                        $sharedType = if ($perm.Link) { $perm.Link.Type } elseif ($perm.Invitation) { "Invitation" } else { "Direct" }
                        $sharedWith = $perm.GrantedToV2?.User?.DisplayName
                        $sharedEmail = $perm.GrantedToV2?.User?.Email
                        $sharedGroup = $perm.GrantedToV2?.Group?.DisplayName
                        $identities = $perm.GrantedToIdentitiesV2 | ForEach-Object {
                            $_.User?.DisplayName, $_.User?.Email, $_.Group?.DisplayName
                        } | Where-Object { $_ } | Select-Object -Unique
                        
                        $null = $results.Add([PSCustomObject]@{
                            DriveName      = $drive.Name
                            FileName       = $file.Name
                            FilePath       = $file.WebUrl
                            FileSize       = $file.Size
                            CreatedDateTime = $file.CreatedDateTime
                            SharedType     = $sharedType
                            SharedWith     = $sharedWith
                            SharedEmail    = $sharedEmail
                            SharedGroup    = $sharedGroup
                            Identities     = ($identities -join "; ")
                            Roles          = ($perm.Roles -join ", ")
                            LastModified   = $file.LastModifiedDateTime
                            PermissionId   = $perm.Id
                            ScanDateTime   = $startTime
                        })
                        
                        if ($VerboseOutput) {
                            Write-Output ("Permission details for $($file.Name): " + ($perm | ConvertTo-Json -Depth 10))
                        }
                    }
                }
            } catch {
                Write-Output "Failed to retrieve permissions for file $($file.Name): $_"
                continue
            }
        }
    } catch {
        Write-Output "Failed to process drive $($drive.Name): $_"
        continue
    }
}

if ($results.Count -eq 0) {
    Write-Output "No sharing permissions found in this site."
} else {
    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Output "Export complete. File saved as $OutputCsv"
    
    # Display summary statistics
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Output ""
    Write-Output "Summary Statistics:"
    Write-Output "• Total files processed: $processedFiles"
    Write-Output "• Total permissions found: $totalPermissions"
    Write-Output "• Document libraries scanned: $($drives.Count)"
    Write-Output "• Processing time: $($duration.TotalMinutes.ToString('F1')) minutes"
    Write-Output "• Scan timestamp: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}
