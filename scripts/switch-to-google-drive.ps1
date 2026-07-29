#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Switches Windows default cloud storage from Microsoft OneDrive to Google Drive.

.DESCRIPTION
    1. Disables OneDrive folder backup (Desktop, Documents, Pictures)
    2. Resets Windows default save locations to local disk
    3. Downloads and installs Google Drive for desktop (if not present)
    4. Opens Google Drive sign-in

.NOTES
    Run in PowerShell as Administrator:
      Set-ExecutionPolicy -Scope Process Bypass -Force
      .\switch-to-google-drive.ps1
#>

$ErrorActionPreference = "Stop"

function Write-Step($msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

function Test-OneDriveRunning {
    return (Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue) -ne $null
}

# ---------------------------------------------------------------------------
# Step 1: Disable OneDrive Known Folder Move (folder backup)
# ---------------------------------------------------------------------------
Write-Step "Disabling OneDrive folder backup"

$kfmKey = "HKCU:\Software\Microsoft\OneDrive\Accounts\Personal"
if (Test-Path $kfmKey) {
    $folders = @("Desktop", "Documents", "Pictures")
    foreach ($folder in $folders) {
        $prop = "KFMFolder$folder"
        if (Get-ItemProperty -Path $kfmKey -Name $prop -ErrorAction SilentlyContinue) {
            Set-ItemProperty -Path $kfmKey -Name $prop -Value 0 -ErrorAction SilentlyContinue
            Write-Host "  Disabled OneDrive backup for: $folder"
        }
    }
} else {
    Write-Host "  OneDrive personal account key not found (may not be configured)."
}

# Stop OneDrive sync temporarily
if (Test-OneDriveRunning) {
    Write-Host "  Stopping OneDrive..."
    Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# ---------------------------------------------------------------------------
# Step 2: Reset default save locations to local disk (C:)
# ---------------------------------------------------------------------------
Write-Step "Setting default save locations to local disk (C:)"

$localDisk = "C:"
$categories = @{
    "DefaultLocation"          = $localDisk
    "DefaultLocationDocuments" = $localDisk
    "DefaultLocationPictures"  = $localDisk
    "DefaultLocationMusic"     = $localDisk
    "DefaultLocationVideos"    = $localDisk
}

$storageKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"
if (-not (Test-Path $storageKey)) {
    New-Item -Path $storageKey -Force | Out-Null
}
foreach ($entry in $categories.GetEnumerator()) {
    Set-ItemProperty -Path $storageKey -Name $entry.Key -Value $entry.Value -Type String
    Write-Host "  $($entry.Key) -> $($entry.Value)"
}

# ---------------------------------------------------------------------------
# Step 3: Move redirected folders back to local paths (if needed)
# ---------------------------------------------------------------------------
Write-Step "Checking user folder locations"

$shellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
$localPaths = @{
  "Personal" = "$env:USERPROFILE\Documents"
  "Desktop"  = "$env:USERPROFILE\Desktop"
  "My Pictures" = "$env:USERPROFILE\Pictures"
}

foreach ($name in $localPaths.Keys) {
    $current = (Get-ItemProperty -Path $shellFolders -Name $name -ErrorAction SilentlyContinue).$name
    if ($current -and $current -match "OneDrive") {
        $target = $localPaths[$name]
        if (-not (Test-Path $target)) { New-Item -ItemType Directory -Path $target -Force | Out-Null }
        Set-ItemProperty -Path $shellFolders -Name $name -Value $target
        Write-Host "  Reset $name from OneDrive to $target"
    }
}

# ---------------------------------------------------------------------------
# Step 4: Install Google Drive for desktop
# ---------------------------------------------------------------------------
Write-Step "Installing Google Drive for desktop"

$gdriveExe = "$env:LOCALAPPDATA\Google\DriveFS\googledrivesync.exe"
$gdriveInstalled = Test-Path $gdriveExe

if (-not $gdriveInstalled) {
    $installerUrl = "https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"
    $installerPath = "$env:TEMP\GoogleDriveSetup.exe"
    Write-Host "  Downloading Google Drive installer..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
    Write-Host "  Running installer (silent)..."
    Start-Process -FilePath $installerPath -ArgumentList "/silent" -Wait
    Start-Sleep -Seconds 5
    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    Write-Host "  Google Drive installed."
} else {
    Write-Host "  Google Drive is already installed."
}

# ---------------------------------------------------------------------------
# Step 5: Launch Google Drive sign-in
# ---------------------------------------------------------------------------
Write-Step "Launching Google Drive sign-in"

$launchPaths = @(
    "$env:LOCALAPPDATA\Google\DriveFS\googledrivesync.exe",
    "${env:ProgramFiles}\Google\Drive File Stream\launch.bat",
    "${env:ProgramFiles(x86)}\Google\Drive File Stream\launch.bat"
)

$launched = $false
foreach ($path in $launchPaths) {
    if (Test-Path $path) {
        Start-Process $path
        $launched = $true
        break
    }
}

if (-not $launched) {
    Write-Host "  Could not auto-launch Google Drive. Open it from the Start menu."
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Host "`n========================================" -ForegroundColor Green
Write-Host " Setup complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host @"

Next steps:
  1. Sign in to Google Drive when the app opens (use your Google account).
  2. Choose 'Stream files' or 'Mirror files' when prompted.
  3. Google Drive will appear under 'This PC' in File Explorer.
  4. Optional: Unlink OneDrive via system tray > Settings > Account > Unlink.

"@
