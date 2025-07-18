<#
.SYNOPSIS
    Exports SharePoint site file permissions to CSV with enhanced features

.DESCRIPTION
    This script connects to Microsoft Graph and exports all file permissions from a SharePoint site to a CSV file.
    It processes all document libraries with detailed permission information, rate limiting, retry logic, and comprehensive error handling.

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

.PARAMETER EnableLogging
    Enable logging to file

.PARAMETER LogFile
    Log file path (default: SPO-Permissions-Log.txt)

.PARAMETER MaxRetries
    Maximum number of retries for failed API calls (default: 3)

.PARAMETER ShowProgress
    Show progress bars during processing (default: true)

.PARAMETER ClientId
    Application (Client) ID for app registration authentication

.PARAMETER TenantId
    Tenant ID for app registration authentication

.PARAMETER ClientSecret
    Client secret for app registration authentication

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -OutputCsv "permissions.csv" -VerboseOutput -EnableLogging

.EXAMPLE
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example" -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -ClientSecret "your-client-secret"

.EXAMPLE
    # Using existing global token
    $global:graphAPIToken = "your-access-token-here"
    .\List-SPO-site-perms.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/example"

.NOTES
    Version: 2.0
    Author: Enhanced SharePoint Permissions Scanner
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
    
    [Parameter(HelpMessage="Enable logging to file")]
    [switch]$EnableLogging,
    
    [Parameter(HelpMessage="Log file path")]
    [string]$LogFile = "SPO-Permissions-Log.txt",
    
    [Parameter(HelpMessage="Maximum number of retries for failed API calls")]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    
    [Parameter(HelpMessage="Show progress bars during processing")]
    [switch]$ShowProgress = $true,
    
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
$scriptVersion = "2.0"
$systemFilePatterns = @("~\$", "\.tmp$", "Forms/", "_vti_", "SitePages/", "SiteAssets/", "Style Library/")

# Rate limiting variables
$script:lastApiCall = Get-Date
$script:minApiInterval = 100 # milliseconds

# Logging function
function Write-Log {
    param(
        [string]$Message, 
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Write-Output $logEntry
    
    if ($EnableLogging) {
        Add-Content -Path $LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
}

# Memory monitoring function
function Get-MemoryUsage {
    $process = Get-Process -Id $PID
    return [Math]::Round($process.WorkingSet64 / 1MB, 2)
}

# Rate limiting function
function Invoke-ThrottledApiCall {
    param([scriptblock]$ApiCall)
    
    $timeSinceLastCall = (Get-Date) - $script:lastApiCall
    if ($timeSinceLastCall.TotalMilliseconds -lt $script:minApiInterval) {
        Start-Sleep -Milliseconds ($script:minApiInterval - $timeSinceLastCall.TotalMilliseconds)
    }
    
    $script:lastApiCall = Get-Date
    return & $ApiCall
}

# Configuration validation function
function Test-Configuration {
    param($SiteUrl, $OutputCsv, $MaxDepth)
    
    Write-Log "Validating configuration..." "INFO"
    
    # Validate site URL format
    if (-not ($SiteUrl -match "^https://.*\.sharepoint\.com/sites/.*")) {
        Write-Log "Site URL doesn't match expected SharePoint format" "WARNING"
    }
    
    # Validate output path
    $outputDir = Split-Path $OutputCsv -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        try {
            New-Item -ItemType Directory -Path $outputDir -Force -ErrorAction Stop | Out-Null
            Write-Log "Created output directory: $outputDir" "INFO"
        } catch {
            Write-Log "Cannot create output directory: $outputDir - $_" "ERROR"
            return $false
        }
    }
    
    # Validate MaxDepth
    if ($MaxDepth -gt 20) {
        Write-Log "MaxDepth > 20 may cause performance issues" "WARNING"
    }
    
    Write-Log "Configuration validation completed" "SUCCESS"
    return $true
}

# Enhanced file inclusion check
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

# Enhanced recursive file retrieval with ArrayList
function Get-FilesRecursively {
    param($DriveId, $ParentId = $null, $CurrentDepth = 0, $MaxDepth = 10)
    
    if ($CurrentDepth -ge $MaxDepth) {
        return [System.Collections.ArrayList]::new()
    }
    
    $allFiles = [System.Collections.ArrayList]::new()
    
    try {
        $items = Invoke-ThrottledApiCall -ApiCall {
            if ($ParentId) {
                Get-MgDriveItem -DriveId $DriveId -DriveItemId $ParentId -ExpandProperty "children" -ErrorAction Stop
            } else {
                Get-MgDriveItem -DriveId $DriveId -Filter "file ne null or folder ne null" -ErrorAction Stop
            }
        }
        
        $children = if ($ParentId) { $items.Children } else { $items }
        
        foreach ($item in $children) {
            if ($item.File) {
                if (Should-IncludeFile -fileName $item.Name -filePath $item.WebUrl) {
                    $null = $allFiles.Add($item)
                }
            } elseif ($item.Folder -and $CurrentDepth -lt $MaxDepth) {
                if ($VerboseOutput) {
                    Write-Log "Scanning folder: $($item.Name) (depth: $($CurrentDepth + 1))" "INFO"
                }
                $subFiles = Get-FilesRecursively -DriveId $DriveId -ParentId $item.Id -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
                $allFiles.AddRange($subFiles)
            }
        }
    } catch {
        Write-Log "Failed to retrieve items at depth $CurrentDepth`: $_" "ERROR"
    }
    
    return $allFiles
}

# Enhanced permission retrieval with retry logic
function Get-FilePermissionsWithRetry {
    param($DriveId, $FileId, $FileName, $MaxRetries = 3)
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            return Invoke-ThrottledApiCall -ApiCall {
                Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $FileId -ErrorAction Stop
            }
        } catch {
            if ($i -eq $MaxRetries) {
                Write-Log "Failed to retrieve permissions for file $FileName after $MaxRetries attempts: $_" "ERROR"
                return $null
            }
            Write-Log "Retry $i/$MaxRetries for file $FileName`: $_" "WARNING"
            Start-Sleep -Seconds 2
        }
    }
}

# Enhanced CSV record creation
function New-CsvRecord {
    param($Site, $Drive, $File, $Permission, $SharedType, $SharedWith, $SharedEmail, $SharedGroup, $Identities, $StartTime)
    
    return [PSCustomObject]@{
        SiteName         = $Site.DisplayName
        SiteUrl          = $Site.WebUrl
        SiteId           = $Site.Id
        DriveName        = $Drive.Name
        DriveId          = $Drive.Id
        FileName         = $File.Name
        FilePath         = $File.WebUrl
        FileExtension    = [System.IO.Path]::GetExtension($File.Name)
        FileSize         = $File.Size
        FileSizeReadable = if ($File.Size) { "{0:N2} MB" -f ($File.Size / 1MB) } else { "N/A" }
        CreatedDateTime  = $File.CreatedDateTime
        ModifiedDateTime = $File.LastModifiedDateTime
        CreatedBy        = $File.CreatedBy?.User?.DisplayName
        ModifiedBy       = $File.LastModifiedBy?.User?.DisplayName
        SharedType       = $SharedType
        SharedWith       = $SharedWith
        SharedEmail      = $SharedEmail
        SharedGroup      = $SharedGroup
        Identities       = $Identities
        Roles            = if ($Permission) { ($Permission.Roles -join ", ") } else { "" }
        LinkScope        = $Permission?.Link?.Scope
        LinkType         = $Permission?.Link?.Type
        PermissionId     = $Permission?.Id
        ScanDateTime     = $StartTime
        ProcessedBy      = $env:USERNAME
        ScriptVersion    = $scriptVersion
    }
}

# Initialize logging
if ($EnableLogging) {
    Write-Log "SharePoint Permissions Scanner v$scriptVersion started" "INFO"
    Write-Log "Parameters: SiteUrl=$SiteUrl, OutputCsv=$OutputCsv, MaxDepth=$MaxDepth" "INFO"
}

# Test configuration
if (-not (Test-Configuration -SiteUrl $SiteUrl -OutputCsv $OutputCsv -MaxDepth $MaxDepth)) {
    Write-Log "Configuration validation failed. Exiting." "ERROR"
    exit 1
}

# Ensure Microsoft Graph module is installed
Write-Log "Checking Microsoft Graph module..." "INFO"
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Installing Microsoft.Graph module..." "INFO"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph

# Enhanced authentication section
try {
    Write-Log "Starting authentication process..." "INFO"
    
    # Check if existing global token is available
    if ($global:graphAPIToken) {
        Write-Log "Found existing global token, will use it for authentication" "INFO"
        $secureToken = ConvertTo-SecureString $global:graphAPIToken -AsPlainText -Force
    }
    
    # Check authentication methods
    if ($ClientId -and $TenantId -and $ClientSecret) {
        # Method 1: App registration with client secret (optionally using existing token)
        Write-Log "Connecting to Microsoft Graph using app registration..." "INFO"
        Write-Log "Client ID: $ClientId" "INFO"
        Write-Log "Tenant ID: $TenantId" "INFO"
        Write-Log "Client Secret: $($ClientSecret.Substring(0, 5))..." "INFO"
        
        # Validate parameters format
        if ([string]::IsNullOrWhiteSpace($ClientId) -or $ClientId.Length -ne 36) {
            Write-Log "ClientId appears invalid. Should be a 36-character GUID" "ERROR"
            exit 1
        }
        
        if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId.Length -ne 36) {
            Write-Log "TenantId appears invalid. Should be a 36-character GUID" "ERROR"
            exit 1
        }
        
        if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
            Write-Log "ClientSecret is empty" "ERROR"
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
            # ... error handling code ...
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

# Add a small delay to ensure cleanup
Start-Sleep -Seconds 1

# Clear any cached authentication state
try {
    Clear-MgContext -ErrorAction SilentlyContinue
} catch {
    # Ignore if command doesn't exist in older versions
}
        
        # Connect using existing token if available, otherwise use client credentials
        try {
            Write-Log "Attempting connection to Microsoft Graph..." "INFO"
            
            if ($global:graphAPIToken) {
                Write-Log "Using existing global token for connection" "INFO"
                Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop
            } else {
                Write-Log "Using client credentials for connection" "INFO"
                $secureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
                $clientCredential = New-Object System.Management.Automation.PSCredential($ClientId, $secureClientSecret)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientCredential -NoWelcome -ErrorAction Stop
            }
            
            Write-Log "Connect-MgGraph command executed successfully" "SUCCESS"
        } catch {
            Write-Log "Connect-MgGraph failed: $($_.Exception.Message)" "ERROR"
            
            # Common error patterns and solutions
            if ($_.Exception.Message -match "AADSTS7000215") {
                Write-Log "Invalid client secret provided - regenerate in Azure Portal" "ERROR"
            } elseif ($_.Exception.Message -match "AADSTS7000222") {
                Write-Log "Client secret has expired - create new secret in Azure Portal" "ERROR"
            } elseif ($_.Exception.Message -match "AADSTS700016") {
                Write-Log "Invalid application ID (ClientId) - verify from app registration" "ERROR"
            } elseif ($_.Exception.Message -match "AADSTS90002") {
                Write-Log "Invalid tenant ID - verify from app registration" "ERROR"
            }
            exit 1
        }
        
        # Verify the connection worked
        Write-Log "Verifying connection context..." "INFO"
        $context = Get-MgContext
        if (-not $context -or -not $context.Account) {
            Write-Log "Authentication failed - no valid context found" "ERROR"
            exit 1
        }
        
        Write-Log "Connection context verified successfully" "SUCCESS"
        Write-Log "Account: $($context.Account)" "INFO"
        Write-Log "AuthType: $($context.AuthType)" "INFO"
        
        # Test basic Graph API access
        Write-Log "Testing basic Graph API access..." "INFO"
        try {
            $profile = Get-MgProfile -ErrorAction Stop
            Write-Log "Basic Graph API access: SUCCESS" "SUCCESS"
        } catch {
            Write-Log "Basic Graph API access: FAILED - $_" "ERROR"
        }
        
        # Test Sites.Read.All permission specifically
        Write-Log "Testing Sites.Read.All permission..." "INFO"
        try {
            $testSites = Get-MgSite -Top 1 -ErrorAction Stop
            Write-Log "Sites.Read.All permission: SUCCESS" "SUCCESS"
        } catch {
            Write-Log "Sites.Read.All permission: FAILED - $_" "ERROR"
            Write-Log "This permission is required for SharePoint access" "ERROR"
            exit 1
        }
        
        Write-Log "App registration authentication fully verified" "SUCCESS"
        
    } else {
        # Method 2: Interactive authentication
        Write-Log "Connecting to Microsoft Graph interactively..." "INFO"
        Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"
        Write-Log "Connected to Microsoft Graph interactively" "SUCCESS"
    }
    
    # Final verification
    $context = Get-MgContext
    if (-not $context -or -not $context.Account) {
        Write-Log "Failed to establish Microsoft Graph connection" "ERROR"
        exit 1
    }
    
    if ($VerboseOutput) {
        Write-Log "Final authentication details:" "INFO"
        Write-Log "Account: $($context.Account)" "INFO"
        Write-Log "Auth Type: $($context.AuthType)" "INFO"
        Write-Log "Scopes: $($context.Scopes -join ', ')" "INFO"
        Write-Log "Tenant ID: $($context.TenantId)" "INFO"
    }
    
} catch {
    Write-Log "Authentication failed: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Site URL validation and parsing
Write-Log "Validating site URL..." "INFO"
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Log "No SharePoint site URL provided" "ERROR"
    exit 1
}

try {
    $uri = [System.Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath.TrimStart("/")
    $hostname = $uri.Hostname
    Write-Log "Site URL parsed successfully: $hostname" "SUCCESS"
} catch {
    Write-Log "Invalid site URL format: $_" "ERROR"
    exit 1
}

# Get Site Object
Write-Log "Searching for SharePoint site..." "INFO"
try {
    # Try to get the site directly by URL
    $site = Get-MgSite -Search "${hostname}:${sitePath}" | Where-Object { $_.WebUrl -eq $SiteUrl }

    if (-not $site) {
        # Fallback: try searching by hostname only
        $site = Get-MgSite -Search $hostname | Where-Object { $_.WebUrl -eq $SiteUrl }
    }

    if (-not $site) {
        Write-Log "Site not found. Please check the URL and try again" "ERROR"
        exit 1
    }
    
    Write-Log "Found site: $($site.DisplayName) (ID: $($site.Id))" "SUCCESS"
} catch {
    if ($_.Exception.Message -match "401|Unauthorized") {
        Write-Log "Access denied (401 Unauthorized) - check app registration permissions" "ERROR"
    } else {
        Write-Log "Failed to retrieve site: $($_.Exception.Message)" "ERROR"
    }
    exit 1
}

# Get Document Libraries (Drives)
Write-Log "Retrieving document libraries..." "INFO"
$drives = Get-MgSiteDrive -SiteId $site.Id
if ($null -eq $drives -or $drives.Count -eq 0) {
    Write-Log "No document libraries found in this site" "WARNING"
    exit 1
}

Write-Log "Found $($drives.Count) document libraries" "SUCCESS"
if ($VerboseOutput) {
    $drives | ForEach-Object { Write-Log "Drive: $($_.Name) (ID: $($_.Id))" "INFO" }
}

# Initialize results collection
$results = [System.Collections.ArrayList]::new()

# Process each drive
foreach ($drive in $drives) {
    Write-Log "Processing drive: $($drive.Name)" "INFO"
    
    try {
        # Get all files recursively
        $allFiles = Get-FilesRecursively -DriveId $drive.Id -MaxDepth $MaxDepth
        
        if ($allFiles.Count -eq 0) {
            Write-Log "No files found in drive: $($drive.Name)" "WARNING"
            continue
        }
        
        Write-Log "Found $($allFiles.Count) files in drive: $($drive.Name)" "INFO"
        
        # Process files with progress tracking
        $fileCount = 0
        $totalFiles = $allFiles.Count
        
        foreach ($file in $allFiles) {
            $fileCount++
            $processedFiles++
            
            # Update progress
            if ($ShowProgress) {
                $percentComplete = [Math]::Round(($fileCount / $totalFiles) * 100, 1)
                Write-Progress -Activity "Processing files in $($drive.Name)" -Status "File $fileCount of $totalFiles ($percentComplete%)" -PercentComplete $percentComplete
            }
            
            if ($VerboseOutput) { 
                Write-Log "Checking file: $($file.Name) ($fileCount/$totalFiles)" "INFO"
            } else {
                # Show progress for non-verbose mode
                if ($fileCount % 10 -eq 0) {
                    Write-Log "Processed $fileCount/$totalFiles files in $($drive.Name)" "INFO"
                }
            }
            
            # Memory cleanup every 100 files
            if ($fileCount % 100 -eq 0) {
                [System.GC]::Collect()
                $memoryMB = Get-MemoryUsage
                if ($VerboseOutput) {
                    Write-Log "Memory usage: $memoryMB MB" "INFO"
                }
            }
            
            # Get file permissions with retry logic
            $permissions = Get-FilePermissionsWithRetry -DriveId $drive.Id -FileId $file.Id -FileName $file.Name -MaxRetries $MaxRetries
            
            if ($null -eq $permissions) {
                continue
            }
            
            if ($permissions.Count -eq 0) {
                # Add entry for files with no explicit permissions
                $csvRecord = New-CsvRecord -Site $site -Drive $drive -File $file -Permission $null -SharedType "None" -SharedWith "No explicit permissions" -SharedEmail "" -SharedGroup "" -Identities "" -StartTime $startTime
                $null = $results.Add($csvRecord)
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
                    
                    $csvRecord = New-CsvRecord -Site $site -Drive $drive -File $file -Permission $perm -SharedType $sharedType -SharedWith $sharedWith -SharedEmail $sharedEmail -SharedGroup $sharedGroup -Identities ($identities -join "; ") -StartTime $startTime
                    $null = $results.Add($csvRecord)
                    
                    if ($VerboseOutput) {
                        Write-Log "Permission details for $($file.Name): $($perm | ConvertTo-Json -Depth 3)" "INFO"
                    }
                }
            }
        }
        
        # Complete progress for this drive
        if ($ShowProgress) {
            Write-Progress -Activity "Processing files in $($drive.Name)" -Completed
        }
        
    } catch {
        Write-Log "Failed to process drive $($drive.Name): $_" "ERROR"
        continue
    }
}

# Export results and display summary
if ($results.Count -eq 0) {
    Write-Log "No sharing permissions found in this site" "WARNING"
} else {
    Write-Log "Exporting results to CSV..." "INFO"
    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log "Export complete. File saved as $OutputCsv" "SUCCESS"
    
    # Display summary statistics
    $endTime = Get-Date
    $duration = $endTime - $startTime
    $finalMemory = Get-MemoryUsage
    
    Write-Log "Processing completed successfully" "SUCCESS"
    Write-Log "=== SUMMARY STATISTICS ===" "INFO"
    Write-Log "Total files processed: $processedFiles" "INFO"
    Write-Log "Total permissions found: $totalPermissions" "INFO"
    Write-Log "Document libraries scanned: $($drives.Count)" "INFO"
    Write-Log "Processing time: $($duration.TotalMinutes.ToString('F1')) minutes" "INFO"
    Write-Log "Peak memory usage: $finalMemory MB" "INFO"
    Write-Log "Scan timestamp: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
    Write-Log "Script version: $scriptVersion" "INFO"
    
    # Show file location
    $fullPath = Resolve-Path $OutputCsv
    Write-Log "Full output path: $fullPath" "INFO"
}

Write-Log "SharePoint Permissions Scanner completed" "SUCCESS"
