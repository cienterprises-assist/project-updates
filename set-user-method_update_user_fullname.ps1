# update_user_fullname.ps1
# Hosted at: [invalid url, do not cite] (GitHub: [invalid url, do not cite])
# Purpose: Updates user full name using Set-LocalUser, with popup notification.

# Ensure admin privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -WindowStyle Hidden
    exit
}

# Set variables
$tempFolder = $env:TEMP
$username = "User"
$fullName = "System Operator (SysOp)"
$logPath = Join-Path $tempFolder "UserFullNameUpdate.log"
$delaySeconds = 15

# Log function
function Write-Log {
    param($Message)
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') IST: $Message" | Out-File -FilePath $logPath -Append -Encoding UTF8
    Write-Host $Message -ForegroundColor Green
}

# Initialize log
Write-Log "Starting user full name update at 05:05 AM IST, July 06, 2025"

# Update full name using Set-LocalUser
try {
    Write-Log "Updating full name for $username to $fullName"
    $user = Get-LocalUser -Name $username -ErrorAction Stop
    if ($user.FullName -eq $fullName) {
        Write-Log "Full name for $username is already $fullName"
        $message = "Full name for $username is already set to $fullName."
    } else {
        Set-LocalUser -Name $username -FullName $fullName -ErrorAction Stop
        Write-Log "Successfully updated full name for $username"
        $message = "Full name for $username updated to $fullName."
    }
    # Show popup
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($message, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    Write-Log "Error updating full name: $_"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Error updating full name for $username: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Delay for visibility
Write-Log "Waiting $delaySeconds seconds"
Start-Sleep -Seconds $delaySeconds

# Cleanup and pause
Write-Log "Cleanup and pausing"
$scriptPath = $PSCommandPath
if ($scriptPath -and (Test-Path $scriptPath)) {
    Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
    Write-Log "Deleted script: $scriptPath"
}
Write-Host "Press Enter to close..." -ForegroundColor Yellow
Read-Host

Write-Log "Script execution completed"