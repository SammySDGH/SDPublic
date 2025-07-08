<#
.SYNOPSIS
  Download all .jsx files from Grafisch/PS Scripts in SDPublic to your local Photoshop Scripts folders.

.DESCRIPTION
  - Auto-detects all "C:\Program Files\Adobe\Adobe Photoshop *\Presets\Scripts" dirs
  - Uses GitHub API (with correct path‐segment encoding) to list & filter .jsx files
  - Downloads each .jsx into every detected Scripts folder

.NOTES
  Run as Administrator (required to write under C:\Program Files).
#>

# —— CONFIGURATION ——
$owner      = 'SammySDGH'
$repo       = 'SDPublic'
$remotePath = 'Grafisch/PS Scripts'

# —— FUNCTIONS ——

function Require-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
       [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
         [Security.Principal.WindowsBuiltinRole] 'Administrator')
    if (-not $isAdmin) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }
}

function Get-DefaultBranch {
    param($owner, $repo)
    $apiUrl = "https://api.github.com/repos/$owner/$repo"
    $meta   = Invoke-RestMethod -Uri $apiUrl -Headers @{
        'User-Agent' = 'PowerShell'
        'Accept'     = 'application/vnd.github.v3+json'
    }
    return $meta.default_branch
}

function Get-RemoteFiles {
    param($owner, $repo, $branch, $path)
    # Split on '/', encode each segment, then rejoin with '/'
    $uriSegments = $path -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }
    $pathUri     = $uriSegments -join '/'
    $url = "https://api.github.com/repos/$owner/$repo/contents/$pathUri?ref=$branch"

    return Invoke-RestMethod -Uri $url -Headers @{
        'User-Agent' = 'PowerShell'
        'Accept'     = 'application/vnd.github.v3+json'
    }
}

function Get-PhotoshopScriptDirs {
    # Find all Adobe Photoshop installs under Program Files
    Get-ChildItem -Path 'C:\Program Files\Adobe' -Directory -Filter 'Adobe Photoshop *' |
      ForEach-Object { Join-Path $_.FullName 'Presets\Scripts' }
}

# —— MAIN ——

Require-Admin

$branch = Get-DefaultBranch -owner $owner -repo $repo
Write-Host "Using branch '$branch'..." 

# Pull directory listing from GitHub
try {
    $items = Get-RemoteFiles -owner $owner -repo $repo -branch $branch -path $remotePath
} catch {
    Write-Error "Failed to retrieve '$remotePath' on branch '$branch': $_"
    exit 1
}

# Filter for .jsx files
$jsxFiles = $items | Where-Object { $_.type -eq 'file' -and $_.name -match '\.jsx$' }
if (-not $jsxFiles) {
    Write-Error "No .jsx files found under '$remotePath' on branch '$branch'."
    exit 1
}

# Find each Photoshop Scripts directory
$destDirs = Get-PhotoshopScriptDirs
if (-not $destDirs) {
    Write-Error "No Photoshop 'Presets\Scripts' folders found under 'C:\Program Files\Adobe'."
    exit 1
}

# Download each .jsx into each Scripts folder
foreach ($dir in $destDirs) {
    if (-not (Test-Path $dir)) {
        Write-Host "Creating folder: $dir"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    foreach ($file in $jsxFiles) {
        $out = Join-Path $dir $file.name
        Write-Host "Downloading $($file.name) → $dir"
        try {
            Invoke-WebRequest -Uri $file.download_url -OutFile $out -UseBasicParsing
        } catch {
            Write-Warning "  → Failed to download $($file.name): $_"
        }
    }
}

Write-Host "`n All .jsx scripts have been installed."
