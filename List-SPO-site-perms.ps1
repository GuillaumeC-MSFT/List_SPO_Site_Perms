param (
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    [string]$OutputCsv = "SharedFilesPermissions.csv",
    [switch]$VerboseOutput
)

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
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

$results = @()
foreach ($drive in $drives) {
    if ($VerboseOutput) { Write-Host "Processing drive: $($drive.Name)" -ForegroundColor Magenta }
    try {
        $items = Get-MgDriveItem -DriveId $drive.Id -Filter "file ne null" -ErrorAction Stop
    } catch {
        Write-Host "Failed to retrieve items for drive $($drive.Name): $_" -ForegroundColor Red
        continue
    }
    $files = $items | Where-Object { $_.File }
    foreach ($file in $files) {
        if ($VerboseOutput) { Write-Host "Checking file: $($file.Name)" -ForegroundColor Gray }
        try {
            $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $file.Id -ErrorAction Stop
        } catch {
            Write-Host "Failed to retrieve permissions for file $($file.Name): $_" -ForegroundColor Red
            continue
        }
        foreach ($perm in $permissions) {
            $sharedType = if ($perm.Link) { $perm.Link.Type } elseif ($perm.Invitation) { "Invitation" } else { "Direct" }
            $sharedWith = $perm.GrantedToV2?.User?.DisplayName
            $sharedEmail = $perm.GrantedToV2?.User?.Email
            $sharedGroup = $perm.GrantedToV2?.Group?.DisplayName
            $identities = $perm.GrantedToIdentitiesV2 | ForEach-Object {
                $_.User?.DisplayName, $_.User?.Email, $_.Group?.DisplayName
            } | Where-Object { $_ } | Select-Object -Unique
            $results += [PSCustomObject]@{
                DriveName      = $drive.Name
                FileName       = $file.Name
                FilePath       = $file.WebUrl
                SharedType     = $sharedType
                SharedWith     = $sharedWith
                SharedEmail    = $sharedEmail
                SharedGroup    = $sharedGroup
                Identities     = ($identities -join "; ")
                Roles          = ($perm.Roles -join ", ")
                LastModified   = $file.LastModifiedDateTime
                PermissionId   = $perm.Id
            }
            if ($VerboseOutput) {
                Write-Host ("Permission details for $($file.Name): " + ($perm | ConvertTo-Json -Depth 10)) -ForegroundColor DarkGray
            }
        }
    }
}

if ($results.Count -eq 0) {
    Write-Host "No sharing permissions found in this site." -ForegroundColor Yellow
} else {
    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Export complete. File saved as $OutputCsv" -ForegroundColor Green
}