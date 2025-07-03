#Requires -RunAsAdministrator
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize logging
$tempFolder = $env:TEMP
$logPath = Join-Path $tempFolder "V2.0toV3.0_Auto_Script.log"
$log = New-Object System.IO.StreamWriter($logPath, $true)
$log.WriteLine("V2.0toV3.0 Auto Script Log - $(Get-Date)")

# Track downloaded files for cleanup
$downloadedFiles = @()

function Log-Message {
    param ($Message)
    $log.WriteLine("$(Get-Date): $Message")
    $log.Flush()
}

function Show-Popup {
    param ($Message, $Title = "Task Completed", $Icon = 64)
    try {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup($Message, 0, $Title, $Icon)
        Log-Message "Displayed popup: $Title - $Message"
    } catch {
        Log-Message "Error displaying popup: $_"
    }
}

function Add-Delay {
    param ($Seconds = 5)
    Log-Message "Adding delay of $Seconds seconds"
    Start-Sleep -Seconds $Seconds
}

function Check-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Log-Message "Error: Script must run as Administrator."
        Show-Popup "This script must be run as Administrator!" "Error" 16
        $log.Close()
        exit
    }
}

function Download-And-Execute-VBS {
    param ($Url, $FileName)
    $path = Join-Path $tempFolder $FileName
    try {
        Log-Message "Downloading $FileName from $Url"
        Invoke-WebRequest -Uri $Url -OutFile $path -ErrorAction Stop
        Log-Message "Downloaded $FileName to $path"
        $downloadedFiles += $path
        if (-not (Test-Path $path)) {
            throw "Downloaded file $FileName not found"
        }
        Log-Message "Executing $FileName"
        Start-Process -FilePath "wscript.exe" -ArgumentList "`"$path`"" -Wait -ErrorAction Stop
        Log-Message "Successfully executed $FileName"
        Show-Popup "Successfully executed $FileName" "VBS Execution"
        Add-Delay
    } catch {
        Log-Message "Error with $FileName: $_"
        Show-Popup "Error with $FileName. Check log at: $logPath" "Error" 48
    }
}

function Download-Wallpaper {
    param ($Url, $FileName)
    $path = Join-Path $tempFolder $FileName
    try {
        Log-Message "Downloading wallpaper $FileName from $Url"
        Invoke-WebRequest -Uri $Url -OutFile $path -ErrorAction Stop
        Log-Message "Downloaded $FileName to $path"
        $downloadedFiles += $path
        Add-Delay
        return $path
    } catch {
        Log-Message "Error downloading $FileName: $_"
        Show-Popup "Error downloading $FileName. Check log at: $logPath" "Error" 48
        return $null
    }
}

function Move-To-Wallpaper-Folder {
    param ($SourcePath, $FileName)
    $destFolder = "C:\Windows\Web\Ci_Walls_MarkX-V3.0"
    $destPath = Join-Path $destFolder $FileName
    try {
        if (-not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Log-Message "Created folder $destFolder"
        }
        if (Test-Path $destPath) {
            Remove-Item -Path $destPath -Force -ErrorAction Stop
            Log-Message "Overwrote existing wallpaper $FileName"
        }
        Move-Item -Path $SourcePath -Destination $destPath -Force -ErrorAction Stop
        Log-Message "Moved $FileName to $destPath"
        Add-Delay
        return $destPath
    } catch {
        Log-Message "Error moving $FileName: $_"
        Show-Popup "Error moving $FileName. Check log at: $logPath" "Error" 48
        return $null
    }
}

function Set-Desktop-Wallpaper {
    param ($Path, $UserSid, $Username)
    try {
        Log-Message "Setting desktop wallpaper for $Username ($UserSid)"
        $regKey = "HKU:\$UserSid\Control Panel\Desktop"
        New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction Stop | Out-Null
        if (-not (Test-Path $regKey)) {
            Log-Message "Registry key $regKey not found for $Username"
            throw "User profile not loaded"
        }
        Set-ItemProperty -Path $regKey -Name Wallpaper -Value $Path -ErrorAction Stop
        Set-ItemProperty -Path $regKey -Name WallpaperStyle -Value 2 -ErrorAction Stop
        Set-ItemProperty -Path $regKey -Name TileWallpaper -Value 0 -ErrorAction Stop
        Start-Process -FilePath "RUNDLL32.EXE" -ArgumentList "user32.dll,UpdatePerUserSystemParameters" -Wait -ErrorAction Stop
        Log-Message "Set desktop wallpaper to $Path for user $Username ($UserSid)"
        Show-Popup "Desktop wallpaper set to $Path for $Username" "Wallpaper Set"
        Add-Delay
    } catch {
        Log-Message "Error setting desktop wallpaper for $Username ($UserSid): $_"
        Show-Popup "Error setting desktop wallpaper for $Username. Check log at: $logPath" "Error" 48
    } finally {
        Remove-PSDrive -Name HKU -ErrorAction SilentlyContinue
    }
}

function Set-LockScreen-Wallpaper {
    param ($Path)
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name LockScreenImage -Value $Path -ErrorAction Stop
        Log-Message "Set lock screen wallpaper to $Path"
        Show-Popup "Lock screen wallpaper set to $Path" "Lock Screen Set"
        Add-Delay
    } catch {
        Log-Message "Error setting lock screen wallpaper: $_"
        Show-Popup "Error setting lock screen wallpaper. Check log at: $logPath" "Error" 48
    }
}

function Configure-LockScreen-Settings {
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name LockScreenImageStatus -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name LockScreenImagePath -Value "C:\Windows\Web\Ci_Walls_MarkX-V3.0\CiOS Lock MARK-X V3.0 Universal ADMIN-USER PANEL.png" -ErrorAction Stop
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name NoLockScreenCamera -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name NoLockScreenSlideshow -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen\Creative" -Name LockImageFlags -Value 1 -ErrorAction Stop
        Log-Message "Configured lock screen settings"
        Show-Popup "Lock screen settings configured" "Lock Screen Settings"
        Add-Delay
    } catch {
        Log-Message "Error configuring lock screen settings: $_"
        Show-Popup "Error configuring lock screen settings. Check log at: $logPath" "Error" 48
    }
}

function Set-ScreenLock {
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name InactivityTimeoutSecs -Value 180 -ErrorAction Stop
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 180 -ErrorAction Stop
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveActive -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaverIsSecure -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path "HKU:\.DEFAULT\Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 180 -ErrorAction Stop
        Set-ItemProperty -Path "HKU:\.DEFAULT\Control Panel\Desktop" -Name ScreenSaveActive -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path "HKU:\.DEFAULT\Control Panel\Desktop" -Name ScreenSaverIsSecure -Value 1 -ErrorAction Stop
        $desktopRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop"
        if (-not (Test-Path $desktopRegPath)) {
            New-Item -Path $desktopRegPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $desktopRegPath -Name ScreenSaverIsSecure -Value 1 -ErrorAction Stop
        Log-Message "Configured screen to lock after 3 minutes with password"
        Show-Popup "Screen lock set to 3 minutes with password" "Screen Lock Set"
        Add-Delay
    } catch {
        Log-Message "Error configuring screen lock: $_"
        Show-Popup "Error configuring screen lock. Check log at: $logPath" "Error" 48
    }
}

function Set-UserFullName {
    param ($Username, $FullName)
    try {
        $user = Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.Name -eq $Username }
        if ($user) {
            $user.FullName = $FullName
            $user.Put() | Out-Null
            Log-Message "Set full name for $Username to $FullName"
            Show-Popup "Full name set for $Username to $FullName" "User Full Name Set"
            Add-Delay
        } else {
            Log-Message "User $Username not found"
            Show-Popup "User $Username not found" "Error" 48
        }
    } catch {
        Log-Message "Error setting full name for $Username: $_"
        Show-Popup "Error setting full name for $Username. Check log at: $logPath" "Error" 48
    }
}

function Delete-Folders {
    param ($FolderNames)
    foreach ($folder in $FolderNames) {
        $path = "C:\Windows\Web\$folder"
        try {
            if (Test-Path $path) {
                $retryCount = 0
                $maxRetries = 3
                while ($retryCount -lt $maxRetries) {
                    try {
                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                        break
                    } catch {
                        $retryCount++
                        Log-Message "Retry $retryCount/$maxRetries for deleting $folder: $_"
                        Start-Sleep -Seconds 2
                    }
                }
                if (Test-Path $path) {
                    throw "Failed to delete $folder after $maxRetries attempts"
                }
                Log-Message "Successfully deleted folder $path"
                Show-Popup "Deleted folder $folder" "Folder Deletion"
            } else {
                Log-Message "Folder $path not found"
                Show-Popup "Folder $folder not found" "Information" 64
            }
            Add-Delay
        } catch {
            Log-Message "Error deleting $folder: $_"
            Show-Popup "Error deleting $folder. Check log at: $logPath" "Error" 48
        }
    }
}

function Cleanup-TempFiles {
    try {
        Log-Message "Starting cleanup of temporary files"
        foreach ($file in $downloadedFiles) {
            if (Test-Path $file) {
                $retryCount = 0
                $maxRetries = 3
                while ($retryCount -lt $maxRetries) {
                    try {
                        Remove-Item -Path $file -Force -ErrorAction Stop
                        Log-Message "Deleted temporary file $file"
                        break
                    } catch {
                        $retryCount++
                        Log-Message "Retry $retryCount/$maxRetries for deleting $file: $_"
                        Start-Sleep -Seconds 2
                    }
                }
                if (Test-Path $file) {
                    Log-Message "Failed to delete $file after $maxRetries attempts"
                }
            }
        }
        Log-Message "Cleanup of temporary files completed"
        Show-Popup "Cleaned up temporary files" "Cleanup Completed"
        Add-Delay
    } catch {
        Log-Message "Error during cleanup: $_"
        Show-Popup "Error during cleanup. Check log at: $logPath" "Error" 48
    }
}

# Check admin privileges
Check-Admin

# Set full names for users
Set-UserFullName -Username "User" -FullName "System Operator (SysOp)"
Set-UserFullName -Username "Agency" -FullName "Admin"
Set-UserFullName -Username "Jarvis" -FullName "SU;CYBERINFORMATICS:JARVIS"

# Execute VBS scripts
$vbsScripts = @(
    @{
        Url = "https://dl.cieverse.com/cios/V3.0 Update/Delete URLAllowlist reg keys chrome and edge.vbs"
        FileName = "Delete URLAllowlist reg keys chrome and edge.vbs"
    },
    @{
        Url = "https://dl.cieverse.com/cios/V3.0 Update/Banks Base Browser/google_chromey_base_V3.0.vbs"
        FileName = "google_chromey_base_V3.0.vbs"
    },
    @{
        Url = "https://dl.cieverse.com/cios/V3.0 Update/Banks Base Browser/edge_latest_base_V3.0.vbs"
        FileName = "edge_latest_base_V3.0.vbs"
    },
    @{
        Url = "https://dl.cieverse.com/cios/V3.0 Update/All Banks Chrome/chrome_latest_icici_bank_V3.0_domain_based.vbs"
        FileName = "chrome_latest_icici_bank_V3.0_domain_based.vbs"
    }
)

foreach ($script in $vbsScripts) {
    Download-And-Execute-VBS -Url $script.Url -FileName $script.FileName
}

# Download and set wallpapers
$userWallpaper = Download-Wallpaper -Url "https://dl.cieverse.com/cios/V3.0 Update/CiE Walls/CiOS Desktop MARK-X User PANEL .png" -FileName "CiOS Desktop MARK-X User PANEL .png"
if ($userWallpaper) {
    $userWallpaperPath = Move-To-Wallpaper-Folder -SourcePath $userWallpaper -FileName "CiOS Desktop MARK-X User PANEL .png"
    if ($userWallpaperPath) {
        Set-Desktop-Wallpaper -Path $userWallpaperPath -UserSid "S-1-5-21-2296551787-2341494431-3366209023-1002" -Username "User"
    }
}

$adminWallpaper = Download-Wallpaper -Url "https://dl.cieverse.com/cios/V3.0 Update/CiE Walls/CiOS Desktop MARK-X Admin PANEL .png" -FileName "CiOS Desktop MARK-X Admin PANEL .png"
if ($adminWallpaper) {
    $adminWallpaperPath = Move-To-Wallpaper-Folder -SourcePath $adminWallpaper -FileName "CiOS Desktop MARK-X Admin PANEL .png"
    if ($adminWallpaperPath) {
        Set-Desktop-Wallpaper -Path $adminWallpaperPath -UserSid "S-1-5-21-2296551787-2341494431-3366209023-1001" -Username "Agency"
        Set-Desktop-Wallpaper -Path $adminWallpaperPath -UserSid "S-1-5-21-2296551787-2341494431-3366209023-500" -Username "Jarvis"
    }
}

$lockWallpaper = Download-Wallpaper -Url "https://dl.cieverse.com/cios/V3.0 Update/CiE Walls/CiOS Lock MARK-X V3.0 Universal ADMIN-USER PANEL.png" -FileName "CiOS Lock MARK-X V3.0 Universal ADMIN-USER PANEL.png"
if ($lockWallpaper) {
    $lockWallpaperPath = Move-To-Wallpaper-Folder -SourcePath $lockWallpaper -FileName "CiOS Lock MARK-X V3.0 Universal ADMIN-USER PANEL.png"
    if ($lockWallpaperPath) {
        Set-LockScreen-Wallpaper -Path $lockWallpaperPath
        Configure-LockScreen-Settings
    }
}

# Configure screen lock after 3 minutes
Set-ScreenLock

# Delete specified folders
$foldersToDelete = @("4K", "Screen", "Wallpaper")
Delete-Folders -FolderNames $foldersToDelete

# Cleanup temporary files
Cleanup-TempFiles

# Finalize
$log.Close()
Show-Popup "All tasks completed. Check log at: $logPath" "Script Completed"
exit
