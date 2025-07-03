# blobfix.ps1
# Purpose: Download and execute blob.vbs using wscript.exe, clean up after VBS popup with robust deletion

# Initialize logging
$logFile = Join-Path $env:TEMP "blobfix_log.txt"
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
}

Write-Log "Script started."

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires administrative privileges. Please run PowerShell as Administrator."
    Write-Log "Error: Script not running as Administrator."
    exit 1
}
Write-Log "Running with administrative privileges."

# Define URLs and paths
$vbsUrl = "https://dl.cieverse.com/cios/MarkX-V2.0 Fixes/blob.vbs"
$tempDir = $env:TEMP
$vbsPath = Join-Path $tempDir "blob.vbs"
$scriptPath = $PSCommandPath

# Download the VBS file to temp directory
Write-Log "Downloading VBS from $vbsUrl to $vbsPath."
try {
    Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsPath -ErrorAction Stop
    Write-Log "VBS downloaded successfully."
}
catch {
    Write-Error "Failed to download VBS file: $_"
    Write-Log "Error downloading VBS: $_"
    exit 1
}

# Verify the VBS file exists and is not empty
if (-not (Test-Path $vbsPath) -or (Get-Item $vbsPath).Length -eq 0) {
    Write-Error "VBS file is missing or empty."
    Write-Log "Error: VBS file is missing or empty."
    exit 1
}
Write-Log "VBS file verified: $vbsPath."

# Execute the VBS file using wscript.exe and wait for completion (including popup)
Write-Log "Executing VBS file: $vbsPath."
try {
    Start-Process -FilePath "wscript.exe" -ArgumentList $vbsPath -Wait -ErrorAction Stop
    Write-Log "VBS executed successfully, popup closed."
}
catch {
    Write-Error "Failed to execute VBS file: $_"
    Write-Log "Error executing VBS: $_"
    exit 1
}

# Clean up: delete VBS and PS1 files after VBS popup is closed
Write-Log "Starting cleanup."
try {
    # Delete blob.vbs if it exists
    if (Test-Path $vbsPath) {
        Remove-Item -Path $vbsPath -Force -ErrorAction Stop
        Write-Log "Deleted VBS file: $vbsPath."
    } else {
        Write-Log "VBS file already deleted or not found: $vbsPath."
    }

    # Delete blobfix.ps1 using a delayed command to avoid self-deletion issues
    if ($scriptPath -and (Test-Path $scriptPath)) {
        $deleteCommand = "Start-Sleep -Milliseconds 1000; Remove-Item -Path `"$scriptPath`" -Force"
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$deleteCommand`"" -NoNewWindow
        Write-Log "Scheduled deletion of script: $scriptPath."
    } else {
        Write-Log "Script path invalid or already deleted: $scriptPath."
    }
}
catch {
    Write-Warning "Cleanup error: $_"
    Write-Log "Warning: Cleanup error: $_"
}

Write-Log "Script completed."
