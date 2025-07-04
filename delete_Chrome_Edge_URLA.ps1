# delete_urlallowlist_chrome_edge.ps1
# Hosted at: https://dl.cieverse.com/ (GitHub: https://github.com/cienterprises-assist/Workstreams)
# Purpose: Downloads and executes a VBS script to delete URLAllowlist registry keys for Chrome and Edge,
#          waits for execution, cleans up the VBS file, and self-deletes this script.

# Ensure the script runs as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator!"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Set up variables
$tempFolder = $env:TEMP
$vbsUrl = "https://dl.cieverse.com/cios/V3.0 Update/Delete URLAllowlist reg keys chrome and edge.vbs"
$vbsFileName = "Delete URLAllowlist reg keys chrome and edge.vbs"
$vbsPath = Join-Path $tempFolder $vbsFileName
$logPath = Join-Path $tempFolder "DeleteBrowserPolicies.log"

# Function to log messages
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    $logMessage | Out-File -FilePath $logPath -Append -Encoding UTF8
}

# Initialize log
Write-Log "Starting delete_urlallowlist_chrome_edge.ps1"

# Download the VBS script
try {
    Write-Log "Downloading VBS script from $vbsUrl"
    Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsPath -ErrorAction Stop
    Write-Log "Successfully downloaded VBS script to $vbsPath"
}
catch {
    Write-Log "Error downloading VBS script: $($_.Exception.Message)"
    Write-Error "Failed to download VBS script: $($_.Exception.Message)"
    exit 1
}

# Verify VBS file exists
if (-not (Test-Path $vbsPath)) {
    Write-Log "VBS script not found at $vbsPath"
    Write-Error "VBS script not found after download"
    exit 1
}

# Execute the VBS script
try {
    Write-Log "Executing VBS script: $vbsPath"
    Start-Process -FilePath "cscript.exe" -ArgumentList "//NoLogo `"$vbsPath`"" -Wait -ErrorAction Stop
    Write-Log "VBS script execution completed"
}
catch {
    Write-Log "Error executing VBS script: $($_.Exception.Message)"
    Write-Error "Failed to execute VBS script: $($_.Exception.Message)"
    # Continue to cleanup
}

# Wait 25 seconds for VBS script popup
Write-Log "Waiting 25 seconds for VBS script popup"
Start-Sleep -Seconds 25

# Delete the VBS script
try {
    if (Test-Path $vbsPath) {
        Remove-Item $vbsPath -Force -ErrorAction Stop
        Write-Log "Successfully deleted VBS script: $vbsPath"
    }
}
catch {
    Write-Log "Error deleting VBS script: $($_.Exception.Message)"
    Write-Error "Failed to delete VBS script: $($_.Exception.Message)"
}

# Self-cleanup: Delete this PowerShell script if stored locally
try {
    $scriptPath = $PSCommandPath
    if ($scriptPath -and (Test-Path $scriptPath)) {
        Write-Log "Deleting PowerShell script: $scriptPath"
        # Create a temporary script to delete this one
        $cleanupScript = Join-Path $tempFolder "cleanup.ps1"
        @"
Start-Sleep -Milliseconds 500
Remove-Item -Path '$scriptPath' -Force -ErrorAction SilentlyContinue
Remove-Item -Path '$cleanupScript' -Force -ErrorAction SilentlyContinue
"@ | Out-File $cleanupScript -Encoding UTF8
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$cleanupScript`"" -NoNewWindow
        Write-Log "Initiated self-cleanup of PowerShell script"
    }
}
catch {
    Write-Log "Error during self-cleanup: $($_.Exception.Message)"
    Write-Error "Failed to delete PowerShell script: $($_.Exception.Message)"
}

# Show final popup
Write-Log "Displaying cleanup successful popup"
$wshell = New-Object -ComObject WScript.Shell
$wshell.Popup("Cleanup successful", 0, "Success", 64)

Write-Log "Script execution completed"
