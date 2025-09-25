#Requires -Version 5.1

<#
.SYNOPSIS
    Final Demo Test Script for PSADT Projects
.DESCRIPTION
    Interactive test script that:
    1. Allows selection of PSADT projects in Final_Demo folder
    2. Runs installation in Windows Sandbox
    3. Waits 2 minutes between install and uninstall
    4. Runs uninstallation
    5. Keeps sandbox open for verification
    
    Perfect for demonstrating complete PSADT install/uninstall workflows.
.EXAMPLE
    .\Test-FinalDemo.ps1
#>

[CmdletBinding()]
param()

function Get-PSADTProjects {
    param([string]$BasePath)
    
    $projects = @()
    $projectFolders = Get-ChildItem -Path $BasePath -Directory
    
    foreach ($folder in $projectFolders) {
        $psadtScript = Join-Path $folder.FullName "Invoke-AppDeployToolkit.ps1"
        if (Test-Path $psadtScript) {
            $projects += [PSCustomObject]@{
                Name = $folder.Name
                Path = $folder.FullName
                ScriptPath = $psadtScript
            }
        }
    }
    
    return $projects
}

function Show-ProjectSelection {
    param([array]$Projects)
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "PSADT Final Demo Test" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Available PSADT Projects:" -ForegroundColor Yellow
    Write-Host ""
    
    for ($i = 0; $i -lt $Projects.Count; $i++) {
        Write-Host "  $($i + 1). $($Projects[$i].Name)" -ForegroundColor White
    }
    
    Write-Host ""
    do {
        $selection = Read-Host "Select project to test (1-$($Projects.Count))"
        
        if ([int]$selection -ge 1 -and [int]$selection -le $Projects.Count) {
            return $Projects[$selection - 1]
        }
        
        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
    } while ($true)
}

function New-CountdownScript {
    param([string]$ProjectPath)
    
    $sandboxPath = Join-Path $ProjectPath "Sandbox"
    if (-not (Test-Path $sandboxPath)) {
        New-Item -ItemType Directory -Path $sandboxPath -Force | Out-Null
    }
    
    $countdownScript = @'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "INSTALLATION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "The application has been installed." -ForegroundColor White
Write-Host "You can now test the application functionality." -ForegroundColor Cyan

# Create a countdown form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Uninstallation Countdown"
$form.Size = New-Object System.Drawing.Size(500,250)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(30,20)
$titleLabel.Size = New-Object System.Drawing.Size(440,30)
$titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
$titleLabel.Text = "Installation Completed Successfully!"
$titleLabel.ForeColor = [System.Drawing.Color]::Green
$titleLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($titleLabel)

# Instructions label
$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Location = New-Object System.Drawing.Point(30,60)
$instructionLabel.Size = New-Object System.Drawing.Size(440,40)
$instructionLabel.Font = New-Object System.Drawing.Font("Arial",10)
$instructionLabel.Text = "Please test the application now.`nUninstallation will begin automatically in 2 minutes."
$instructionLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($instructionLabel)

# Countdown label
$countdownLabel = New-Object System.Windows.Forms.Label
$countdownLabel.Location = New-Object System.Drawing.Point(30,120)
$countdownLabel.Size = New-Object System.Drawing.Size(440,40)
$countdownLabel.Font = New-Object System.Drawing.Font("Arial",16,[System.Drawing.FontStyle]::Bold)
$countdownLabel.Text = "Time remaining: 02:00"
$countdownLabel.ForeColor = [System.Drawing.Color]::Blue
$countdownLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($countdownLabel)

# Skip button
$skipButton = New-Object System.Windows.Forms.Button
$skipButton.Location = New-Object System.Drawing.Point(200,170)
$skipButton.Size = New-Object System.Drawing.Size(100,30)
$skipButton.Text = "Skip Wait"
$skipButton.Font = New-Object System.Drawing.Font("Arial",10)
$form.Controls.Add($skipButton)

# Timer for countdown
$timer = New-Object System.Windows.Forms.Timer
$secondsRemaining = 120  # 2 minutes
$timer.Interval = 1000  # 1 second

$timer.Add_Tick({
    $script:secondsRemaining--
    
    $minutes = [Math]::Floor($script:secondsRemaining / 60)
    $seconds = $script:secondsRemaining % 60
    
    $countdownLabel.Text = "Time remaining: $($minutes.ToString("00")):$($seconds.ToString("00"))"
    
    if ($script:secondsRemaining -le 0) {
        $timer.Stop()
        $form.Close()
    }
    
    # Change color as time runs out
    if ($script:secondsRemaining -le 30) {
        $countdownLabel.ForeColor = [System.Drawing.Color]::Red
    } elseif ($script:secondsRemaining -le 60) {
        $countdownLabel.ForeColor = [System.Drawing.Color]::Orange
    }
})

# Skip button click event
$skipButton.Add_Click({
    $timer.Stop()
    $form.Close()
})

# Start the timer and show the form
$timer.Start()
$form.ShowDialog()

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "STARTING UNINSTALLATION PROCESS" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
'@

    $countdownPath = Join-Path $sandboxPath "Countdown.ps1"
    Set-Content -Path $countdownPath -Value $countdownScript -Encoding UTF8
    
    return $countdownPath
}

# Main execution
try {
    $basePath = $PSScriptRoot
    if (-not $basePath) {
        $basePath = Get-Location
    }
    
    Write-Host "PSADT Final Demo Test Script" -ForegroundColor Cyan
    Write-Host "============================" -ForegroundColor Cyan
    
    # Find PSADT projects
    $projects = Get-PSADTProjects -BasePath $basePath
    
    if ($projects.Count -eq 0) {
        Write-Host "No PSADT projects found in: $basePath" -ForegroundColor Red
        Write-Host "Please ensure you have PSADT projects with Invoke-AppDeployToolkit.ps1 files." -ForegroundColor Yellow
        exit 1
    }
    
    # Let user select project
    $selectedProject = Show-ProjectSelection -Projects $projects
    
    Write-Host "`nSelected project: $($selectedProject.Name)" -ForegroundColor Green
    Write-Host "Project path: $($selectedProject.Path)" -ForegroundColor Gray
    
    # Create countdown script only
    Write-Host "`nCreating countdown script..." -ForegroundColor Yellow
    
    $countdownPath = New-CountdownScript -ProjectPath $selectedProject.Path
    
    Write-Host "✓ Countdown script: $countdownPath" -ForegroundColor Green
    
    # Create sandbox configuration
    $sandboxFolder = Join-Path $selectedProject.Path "Sandbox"
    $sandboxConfigFile = Join-Path $sandboxFolder "FinalDemo.wsb"
    
    $sandboxConfigContent = @"
<Configuration>
    <VGpu>Disable</VGpu>
    <Networking>Enable</Networking>
    <MappedFolders>
        <MappedFolder>
            <HostFolder>$($selectedProject.Path)</HostFolder>
            <SandboxFolder>C:\PSADT</SandboxFolder>
            <ReadOnly>false</ReadOnly>
        </MappedFolder>
    </MappedFolders>
    <LogonCommand>
        <Command>powershell.exe -NoExit -ExecutionPolicy Bypass -Command &quot;&amp; { C:\PSADT\Invoke-AppDeployToolkit.ps1; C:\PSADT\Sandbox\Countdown.ps1; C:\PSADT\Invoke-AppDeployToolkit.ps1 -DeploymentType Uninstall }&quot;</Command>
    </LogonCommand>
</Configuration>
"@

    Set-Content -Path $sandboxConfigFile -Value $sandboxConfigContent -Encoding UTF8
    Write-Host "✓ Sandbox config: $sandboxConfigFile" -ForegroundColor Green
    
    # Launch Windows Sandbox
    Write-Host "`nLaunching Windows Sandbox for Final Demo..." -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "The sandbox will:" -ForegroundColor White
    Write-Host "1. Install the application" -ForegroundColor Green
    Write-Host "2. Show a 2-minute countdown for testing" -ForegroundColor Yellow
    Write-Host "3. Uninstall the application" -ForegroundColor Red
    Write-Host "4. Keep the sandbox open for verification" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    
    Start-Process -FilePath "WindowsSandbox.exe" -ArgumentList $sandboxConfigFile
    
    Write-Host "`n✓ Final demo sandbox launched successfully!" -ForegroundColor Green
    Write-Host "Monitor the sandbox for the complete install/uninstall cycle." -ForegroundColor White
    
} catch {
    Write-Host "`n❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}