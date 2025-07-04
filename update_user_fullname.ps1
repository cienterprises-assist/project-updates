# update_user_fullname.ps1
# Hosted at: https://dl.cieverse.com/ (GitHub: https://github.com/cienterprises-assist/Workstreams)
# Purpose: Downloads and executes update_user_fullname.vbs to update user full name, with cleanup.
# Note: Ensure https://dl.cieverse.com/cios/V3.0 update/update_user_fullname.vbs is accessible; adjust URL if 404 occurs.

# Ensure the script runs as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator!"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -WindowStyle Hidden
    exit
}

# Set up variables
$tempFolder = $env:TEMP
$vbsUrl = "https://dl.cieverse.com/cios/V3.0 Update/update_user_fullname.vbs"
$vbsFileName = "update_user_fullname.vbs"
$vbsPath = Join-Path $tempFolder $vbsFileName
$delaySeconds = 15
$logPath = Join-Path $tempFolder "UserFullNameAutomation.log"

# Function to log messages
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') IST: $Message"
    $logMessage | Out-File -FilePath $logPath -Append -Encoding UTF8
}

# Initialize log
Write-Log "Starting user full name automation script at 04:00 AM IST, July 06, 2025"

# Download the VBS script
try {
    Write-Log "Downloading ${vbsFileName} from $vbsUrl"
    Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsPath -ErrorAction Stop
    Write-Log "Successfully downloaded ${vbsFileName} to $vbsPath"
}
catch {
    Write-Log "Error downloading ${vbsFileName}: $($_.Exception.Message)"
    Write-Error "Failed to download ${vbsFileName}: $($_.Exception.Message)"
    exit 1
}

# Verify VBS file exists
if (-not (Test-Path $vbsPath)) {
    Write-Log "VBS script not found at $vbsPath"
    Write-Error "VBS script not found after download: ${vbsFileName}"
    exit 1
}

# Execute the VBS script with wscript
try {
    Write-Log "Executing VBS script: $vbsPath"
    Start-Process -FilePath "wscript.exe" -ArgumentList "`"$vbsPath`"" -Wait -WindowStyle Hidden -ErrorAction Stop
    Write-Log "VBS script execution completed: ${vbsFileName}"
}
catch {
    Write-Log "Error executing ${vbsFileName}: $($_.Exception.Message)"
    Write-Error "Failed to execute ${vbsFileName}: $($_.Exception.Message)"
    # Continue to cleanup despite error
}

# Delay to allow popup
Write-Log "Waiting $delaySeconds seconds for ${vbsFileName} popup"
Start-Sleep -Seconds $delaySeconds

# Cleanup process
Write-Log "Starting cleanup process"
try {
    if (Test-Path $vbsPath) {
        Remove-Item $vbsPath -Force -ErrorAction Stop
        Write-Log "Successfully deleted $vbsPath"
    }

    # Self-cleanup: Delete this PowerShell script if stored locally
    $scriptPath = $PSCommandPath
    if ($scriptPath -and (Test-Path $scriptPath)) {
        Write-Log "Deleting PowerShell script: $scriptPath"
        $cleanupScript = Join-Path $tempFolder "cleanup.ps1"
        @"
Start-Sleep -Milliseconds 500
Remove-Item -Path '$scriptPath' -Force -ErrorAction SilentlyContinue
Remove-Item -Path '$cleanupScript' -Force -ErrorAction SilentlyContinue
"@ | Out-File $cleanupScript -Encoding UTF8
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$cleanupScript`"" -WindowStyle Hidden -NoNewWindow
        Write-Log "Initiated self-cleanup of PowerShell script"
    }
}
catch {
    Write-Log "Error during cleanup: $($_.Exception.Message)"
    Write-Error "Failed to cleanup: $($_.Exception.Message)"
}

# Show final popup
Write-Log "Displaying cleanup successful popup"
$wshell = New-Object -ComObject WScript.Shell
$wshell.Popup("USER_FULLNAME_UPDATE_CLEANUP Completed", 0, "Success", 64)

Write-Log "Script execution completed"
