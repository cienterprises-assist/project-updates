# rediffmail-enable-disable.ps1
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Host "Run as Administrator." -ForegroundColor Red; exit 1 }

# Import Windows Forms for popups
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$tempDir = $env:TEMP
$scriptPath = $MyInvocation.MyCommand.Path
$firstExeUrl = "https://dl.cieverse.com/cios/rediffmail-login-permit/rediffmail-acc-enable.vbs"
$secondExeUrl = "https://dl.cieverse.com/cios/rediffmail-login-permit/rediffmail-acc-disable.vbs"
$firstExePath = Join-Path $tempDir "rediffmail-acc-enable.vbs"
$secondExePath = Join-Path $tempDir "rediffmail-acc-disable.vbs"

try {
    # Download first .exe
    Write-Host "Downloading first executable..."
    Invoke-WebRequest -Uri $firstExeUrl -OutFile $firstExePath -ErrorAction Stop

    # Run first .exe (adds registry, shows success popup)
    Write-Host "Executing first executable..."
    $firstExeProcess = Start-Process -FilePath $firstExePath -PassThru -Wait

    # Show disclaimer popup
    $disclaimerResult = [System.Windows.Forms.MessageBox]::Show(
        "Please click 'OK' on the success popup from the first executable to proceed.",
        "Disclaimer",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    # 210-second timer popup with instructions (no OK/close buttons)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Instructions to follow:"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(360, 150)
    $label.Text = "Instructions:`n1. Open Chrome browser, close it, open it again, and close it 3 times to take effect.`n2. Log in to your Rediff Mail by using the link https://mail.rediff.com/cgi-bin/login.cgi, then close Chrome browser.`n3. Wait for the timer to finish the process.`n`nNote: If you encounter any issues, please contact your technical team."
    $form.Controls.Add($label)

    $timerLabel = New-Object System.Windows.Forms.Label
    $timerLabel.Location = New-Object System.Drawing.Point(10, 160)
    $timerLabel.Size = New-Object System.Drawing.Size(360, 30)
    $form.Controls.Add($timerLabel)

    $timer = New-Object System.Windows.Forms.Timer
    $secondsRemaining = 210 # 3.5 minutes
    $timer.Interval = 1000 # Update every second
    $timer.Add_Tick({
        $script:secondsRemaining--
        $minutes = [math]::Floor($secondsRemaining / 60)
        $seconds = $secondsRemaining % 60
        $timerLabel.Text = "Time Remaining: $minutes minutes $seconds seconds"
        if ($secondsRemaining -le 0) {
            $timer.Stop()
            $form.Close()
        }
    })
    $timer.Start()
    $form.ShowDialog() | Out-Null

    # Download second .exe
    Write-Host "Downloading second executable..."
    Invoke-WebRequest -Uri $secondExeUrl -OutFile $secondExePath -ErrorAction Stop

    # Run second .exe (deletes registry, shows auto-closing success popup)
    Write-Host "Executing second executable..."
    Start-Process -FilePath $secondExePath -Wait

    # Final popup for Rediff Mail login confirmation
    $loginResult = [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Click OK if you logged in successfully to Rediff Mail.",
        "64",
        "Rediff Mail Login Confirmation"
    )
    if ($loginResult -ne "OK") {
        Write-Host "User did not confirm successful login." -ForegroundColor Yellow
    }

    # Restart prompt
    $restartResult = [System.Windows.Forms.MessageBox]::Show(
        "The process is complete. Would you like to restart now?",
        "Restart Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Restarting now..."
        Restart-Computer -Force
    } else {
        Write-Host "Please restart your computer later to complete the process."
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show(
        "An error occurred: $($_.Exception.Message)`nContact your technical team.",
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
finally {
    # Cleanup: Remove .exe files and script
    Write-Host "Cleaning up temporary files..."
    if (Test-Path $firstExePath) { Remove-Item $firstExePath -Force }
    if (Test-Path $secondExePath) { Remove-Item $secondExePath -Force }
    if ($scriptPath -and (Test-Path $scriptPath)) { Remove-Item $scriptPath -Force }
}
