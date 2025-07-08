<#
.SYNOPSIS
  Download all .jsx files from Grafisch/PS Scripts in SDPublic to your local Photoshop Scripts folders.

.DESCRIPTION
  - Auto-detects all "C:\Program Files\Adobe\Adobe Photoshop *\Presets\Scripts" dirs
  - Uses GitHub API to list & filter .jsx files in Grafisch/PS Scripts on the default branch
  - Downloads each .jsx into every detected Scripts folder
#>

# ——— CONFIG ———
$owner = 'SammySDGH'
$repo  = 'SDPublic'
$remotePath = 'Grafisch/PS Scripts'

# ——— FUNCS ———
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
    $api = "https://api.github.com/repos/$owner/$repo"
    return (Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent'='PS' }).default_branch
}

function Get-RemoteFiles {
    param($owner, $repo, $branch, $path)
    $escaped = [uri]::EscapeDataString($path)
    $url = "https://api.github.com/repos/$owner/$repo/contents/$escaped?ref=$branch"
    return Invoke-RestMethod -Uri $url -Headers @{ 'User-Agent' = 'PS' }
}

function Get-DestFolders {
    # Finds all Photoshop installs under Program Files
    Get-ChildItem -Path 'C:\Program Files\Adobe' `
                  -Filter 'Adobe Photoshop *' -Directory |
      ForEach-Object { Join-Path $_.FullName 'Presets\Scripts' }
}

# ——— MAIN ———
Require-Admin

$branch = Get-DefaultBranch -owner $owner -repo $repo
Write-Host "Using branch '$branch'..."

$items = Get-RemoteFiles -owner $owner -repo $repo `
                        -branch $branch -path $remotePath

# Filter for .jsx files
$jsxFiles = $items | Where-Object { $_.type -eq 'file' -and $_.name -like '*.jsx' }
if (-not $jsxFiles) {
    Write-Error "No .jsx files found under '$remotePath' on branch '$branch'."
    exit 1
}

$destDirs = Get-DestFolders
if (-not $destDirs) {
    Write-Error "No Photoshop 'Presets\Scripts' folders found under 'C:\Program Files\Adobe'."
    exit 1
}

foreach ($dir in $destDirs) {
    if (-not (Test-Path $dir)) {
        Write-Host "Creating folder: $dir"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    foreach ($file in $jsxFiles) {
        $out = Join-Path $dir $file.name
        Write-Host "Downloading $($file.name) → $dir"
        Invoke-WebRequest -Uri $file.download_url -OutFile $out -UseBasicParsing
    }
}

Write-Host "`n All .jsx scripts have been installed."
