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
                    Write-Host "  Scanning folder: $($item.Name) (depth: $($CurrentDepth + 1))" -ForegroundColor DarkYellow
                }
                $allFiles += Get-FilesRecursively -DriveId $DriveId -ParentId $item.Id -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            }
        }
    } catch {
        Write-Host "Failed to retrieve items at depth $CurrentDepth`: $_" -ForegroundColor Red
    }
    
    return $allFiles
}

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
try {
    # Check authentication methods in order of preference
    if ($global:graphAPIToken) {
        # Method 1: Use existing global token
        Write-Host "Connecting to Microsoft Graph using existing token..." -ForegroundColor Yellow
        $secureToken = ConvertTo-SecureString $global:graphAPIToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
        Write-Host "Connected to Microsoft Graph using existing token." -ForegroundColor Green
    } elseif ($ClientId -and $TenantId -and $ClientSecret) {
        # Method 2: App registration with client secret
        Write-Host "Connecting to Microsoft Graph using app registration..." -ForegroundColor Yellow
        
        # Convert client secret to secure string
        $secureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $clientCredential = New-Object System.Management.Automation.PSCredential($ClientId, $secureClientSecret)
        
        # Connect using client credentials
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientCredential -NoWelcome
     
        # Verify the connection worked
        $context = Get-MgContext
        if (-not $context -or -not $context.Account) {
            Write-Host "Failed to authenticate with app registration. Please check your ClientId, TenantId, and ClientSecret." -ForegroundColor Red
            Write-Host "Make sure the app registration has the required permissions (Sites.Read.All, Files.Read.All, User.Read.All) and admin consent is granted." -ForegroundColor Yellow
            exit 1
        }
        
        Write-Host "Connected to Microsoft Graph using app registration." -ForegroundColor Green
    } else {
        # Method 3: Interactive authentication (existing behavior)
        Write-Host "Connecting to Microsoft Graph interactively..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
   # Final verification that we're connected
    $context = Get-MgContext
    if (-not $context -or -not $context.Account) {
        Write-Host "Failed to establish Microsoft Graph connection." -ForegroundColor Red
        exit 1
    }
    
    if ($VerboseOutput) {
        Write-Host "Graph context: Account=$($context.Account), Scopes=$($context.Scopes -join ', ')" -ForegroundColor DarkGray
    }
    
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit 1
}

# Validate Site URL
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "No SharePoint site URL provided. Exiting." -ForegroundColor Red
    exit 1
}

try {
    $uri = [System.Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath.TrimStart("/")
} catch {
    Write-Host "Invalid site URL format. Exiting." -ForegroundColor Red
    exit 1
}

$hostname = $uri.Hostname

# Get Site Object
# Try to get the site directly by URL
$site = Get-MgSite -Search "${hostname}:${sitePath}" | Where-Object { $_.WebUrl -eq $SiteUrl }

if (-not $site) {
    # Fallback: try searching by hostname only
    $site = Get-MgSite -Search $hostname | Where-Object { $_.WebUrl -eq $SiteUrl }
}

if (-not $site) {
    Write-Host "Site not found. Please check the URL and try again." -ForegroundColor Red
    exit 1
}
$siteId = $site.Id
Write-Host "Found site: $($site.DisplayName) (ID: $siteId)" -ForegroundColor Green

# Get Document Libraries (Drives)
$drives = Get-MgSiteDrive -SiteId $siteId
if ($null -eq $drives -or $drives.Count -eq 0) {
    Write-Host "No document libraries found in this site." -ForegroundColor Red
    exit 1
}
if ($VerboseOutput) {
    Write-Host "Document libraries found:" -ForegroundColor Cyan
    $drives | ForEach-Object { Write-Host "Drive: $($_.Name) (ID: $($_.Id))" -ForegroundColor Cyan }
}

# Use ArrayList for better performance
$results = [System.Collections.ArrayList]::new()

foreach ($drive in $drives) {
    if ($VerboseOutput) { Write-Host "Processing drive: $($drive.Name)" -ForegroundColor Magenta }
    
    try {
        # Get all files recursively
        $allFiles = Get-FilesRecursively -DriveId $drive.Id -MaxDepth $MaxDepth
        
        if ($allFiles.Count -eq 0) {
            Write-Host "No files found in drive: $($drive.Name)" -ForegroundColor Yellow
            continue
        }
        
        $fileCount = 0
        foreach ($file in $allFiles) {
            $fileCount++
            $processedFiles++
            
            if ($VerboseOutput) { 
                Write-Host "Checking file: $($file.Name) ($fileCount/$($allFiles.Count))" -ForegroundColor Gray 
            } else {
                # Show progress for non-verbose mode
                if ($fileCount % 10 -eq 0) {
                    Write-Host "Processed $fileCount/$($allFiles.Count) files in $($drive.Name)..." -ForegroundColor Green
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
                            Write-Host ("Permission details for $($file.Name): " + ($perm | ConvertTo-Json -Depth 10)) -ForegroundColor DarkGray
                        }
                    }
                }
            } catch {
                Write-Host "Failed to retrieve permissions for file $($file.Name): $_" -ForegroundColor Red
                continue
            }
        }
    } catch {
        Write-Host "Failed to process drive $($drive.Name): $_" -ForegroundColor Red
        continue
    }
}

if ($results.Count -eq 0) {
    Write-Host "No sharing permissions found in this site." -ForegroundColor Yellow
} else {
    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Export complete. File saved as $OutputCsv" -ForegroundColor Green
    
    # Display summary statistics
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Host "`n📊 Summary Statistics:" -ForegroundColor Cyan
    Write-Host "• Total files processed: $processedFiles" -ForegroundColor White
    Write-Host "• Total permissions found: $totalPermissions" -ForegroundColor White
    Write-Host "• Document libraries scanned: $($drives.Count)" -ForegroundColor White
    Write-Host "• Processing time: $($duration.TotalMinutes.ToString('F1')) minutes" -ForegroundColor White
    Write-Host "• Scan timestamp: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
}
