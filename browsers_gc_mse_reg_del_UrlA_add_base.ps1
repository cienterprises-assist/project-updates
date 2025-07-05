# browsers_gc_mse_reg_del_UrlA_add_base.ps1
# Hosted at: https://dl.cieverse.com/ (GitHub: https://github.com/cienterprises-assist/Workstreams)
# Purpose: Automates deletion and configuration of Chrome/Edge URLAllowlist and base policies, with cleanup.

# Ensure the script runs as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator!"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -WindowStyle Hidden
    exit
}

# Set up variables
$tempFolder = $env:TEMP
$delaySeconds = 15
$logPath = Join-Path $tempFolder "BrowserPolicyAutomation.log"

# Function to log messages
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    $logMessage | Out-File -FilePath $logPath -Append -Encoding UTF8
}

# Initialize log
Write-Log "Starting browser policy automation script"

# Define VBS file details
$vbsFiles = @(
    @{Url = "https://dl.cieverse.com/cios/V3.0 Update/Delete URLAllowlist reg keys chrome and edge.vbs"; FileName = "Delete URLAllowlist reg keys chrome and edge.vbs" },
    @{Url = "https://dl.cieverse.com/cios/V3.0 Update/Banks Base Browser/google_chromey_base_V3.0.vbs"; FileName = "google_chromey_base_V3.0.vbs" },
    @{Url = "https://dl.cieverse.com/cios/V3.0 Update/Banks Base Browser/edge_latest_base_V3.0.vbs"; FileName = "edge_latest_base_V3.0.vbs" }
)

# Process each VBS file
foreach ($vbs in $vbsFiles) {
    $vbsPath = Join-Path $tempFolder $vbs.FileName

    # Download the VBS script
    try {
        Write-Log "Downloading $($vbs.FileName) from $($vbs.Url)"
        Invoke-WebRequest -Uri $vbs.Url -OutFile $vbsPath -ErrorAction Stop
        Write-Log "Successfully downloaded $($vbs.FileName) to $vbsPath"
    }
    catch {
        Write-Log "Error downloading $($vbs.FileName): $($_.Exception.Message)"
        Write-Error "Failed to download $($vbs.FileName): $($_.Exception.Message)"
        exit 1
    }

    # Verify VBS file exists
    if (-not (Test-Path $vbsPath)) {
        Write-Log "VBS script not found at $vbsPath"
        Write-Error "VBS script not found after download: $($vbs.FileName)"
        exit 1
    }

    # Execute the VBS script
    try {
        Write-Log "Executing VBS script: $vbsPath"
        Start-Process -FilePath "cscript.exe" -ArgumentList "//NoLogo `"$vbsPath`"" -Wait -WindowStyle Hidden -ErrorAction Stop
        Write-Log "VBS script execution completed: $($vbs.FileName)"
    }
    catch {
        Write-Log "Error executing $($vbs.FileName): $($_.Exception.Message)"
        Write-Error "Failed to execute $($vbs.FileName): $($_.Exception.Message)"
        # Continue to next step despite error
    }

    # Delay to allow popup
    Write-Log "Waiting $delaySeconds seconds for $($vbs.FileName) popup"
    Start-Sleep -Seconds $delaySeconds
}

# Cleanup process
Write-Log "Starting cleanup process"
try {
    # Remove all downloaded VBS files
    foreach ($vbs in $vbsFiles) {
        $vbsPath = Join-Path $tempFolder $vbs.FileName
        if (Test-Path $vbsPath) {
            Remove-Item $vbsPath -Force -ErrorAction Stop
            Write-Log "Successfully deleted $vbsPath"
        }
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
$wshell.Popup("URLA_OLD_RM_URLA_BASE_MSE_GC_ADD_CLEANUP Completed", 0, "Success", 64)

Write-Log "Script execution completed"
