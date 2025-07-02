# updateic.ps1
# Ensure script runs with elevated privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrative privileges. Please run PowerShell as Administrator." -ForegroundColor Red
    exit 1
}

# Set variables
$baseUrl = "https://dl.cieverse.com/cios/Zmail-login-en-dis" # Adjust to your folder
$tempDir = $env:TEMP
$scriptPath = $MyInvocation.MyCommand.Path

try {
    # Prompt user for the filename with extension
    $fileName = Read-Host "Enter the name of the file with extension (e.g., ic-update-1.2.3.exe or script.vbs)"
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        Write-Host "Filename cannot be empty." -ForegroundColor Red
        exit 1
    }

    # Construct download URL and output path
    $downloadUrl = "$baseUrl/$fileName"
    $filePath = Join-Path $tempDir $fileName

    # Download the file
    Write-Host "Downloading $fileName..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $filePath

    # Execute the file based on extension
    Write-Host "Executing $fileName..."
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    switch ($extension) {
        ".exe" { Start-Process -FilePath $filePath -Wait }
        ".vbs" { Start-Process -FilePath "cscript.exe" -ArgumentList "//NoLogo `"$filePath`"" -Wait }
        ".ps1" { Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$filePath`"" -Wait }
        default { 
            Write-Host "Unsupported file type: $extension" -ForegroundColor Red
            exit 1
        }
    }

    Write-Host "Execution completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Clean up: delete the downloaded file and the script itself
    if (Test-Path $filePath) {
        Remove-Item $filePath -Force
    }
    if ($scriptPath -and (Test-Path $scriptPath)) {
        Remove-Item $scriptPath -Force
    }
}
