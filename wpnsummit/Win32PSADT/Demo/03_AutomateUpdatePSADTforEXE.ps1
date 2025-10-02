#Requires -Version 5.1

<#
.SYNOPSIS
    Automatically updates Invoke-AppDeployToolkit.ps1 for EXE installers based on YAML configuration
.DESCRIPTION
    Scans the Files folder for YAML files, parses installer information, and updates the PSADT deployment script
    to handle EXE installers with proper parameters extracted from the YAML manifest
.PARAMETER ProjectPath
    Path to the PSADT project folder containing Invoke-AppDeployToolkit.ps1
.PARAMETER FilesPath
    Path to the Files folder containing the installer and YAML files (defaults to ProjectPath\Files)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,
    
    [Parameter(Mandatory = $false)]
    [string]$FilesPath
)

# Set default FilesPath if not provided
if (-not $FilesPath) {
    $FilesPath = Join-Path $ProjectPath "Files"
}

# Variables
$DeploymentScriptPath = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"

function Get-YAMLInstallerInfo {
    param([string]$YAMLFilePath)
    
    try {
        Write-Host "Parsing YAML file: $YAMLFilePath" -ForegroundColor Yellow
        
        $yamlContent = Get-Content $YAMLFilePath -Raw
        $installerInfo = @{
            InstallerType = $null
            SilentArgs = $null
            SilentWithProgressArgs = $null
            InstallLocationArgs = $null
            ProductCode = $null
            PackageName = $null
            Publisher = $null
            PackageVersion = $null
            Architecture = $null
        }
        
        # Parse basic package info
        if ($yamlContent -match 'PackageName:\s*(.+)') {
            $installerInfo.PackageName = $matches[1].Trim()
        }
        if ($yamlContent -match 'Publisher:\s*(.+)') {
            $installerInfo.Publisher = $matches[1].Trim()
        }
        if ($yamlContent -match 'PackageVersion:\s*(.+)') {
            $installerInfo.PackageVersion = $matches[1].Trim()
        }
        
        # Parse installer section using more specific regex patterns
        # Get installer type
        if ($yamlContent -match 'InstallerType:\s*(.+)') {
            $installerInfo.InstallerType = $matches[1].Trim()
        }
        
        # Get architecture
        if ($yamlContent -match 'Architecture:\s*(.+)') {
            $installerInfo.Architecture = $matches[1].Trim()
        }
        
        # Get product code
        if ($yamlContent -match 'ProductCode:\s*(.+)') {
            $installerInfo.ProductCode = $matches[1].Trim()
        }
        
        # Parse installer switches - look for the specific indented pattern
        if ($yamlContent -match 'InstallerSwitches:\s*\n\s+Silent:\s*(.+)') {
            $installerInfo.SilentArgs = $matches[1].Trim()
        }
        if ($yamlContent -match 'InstallerSwitches:.*?\n\s+SilentWithProgress:\s*(.+)') {
            $installerInfo.SilentWithProgressArgs = $matches[1].Trim()
        }
        if ($yamlContent -match 'InstallerSwitches:.*?\n.*?InstallLocation:\s*(.+)') {
            $installerInfo.InstallLocationArgs = $matches[1].Trim()
        }
        
        return $installerInfo
    }
    catch {
        Write-Error "Failed to parse YAML file: $($_.Exception.Message)"
        return $null
    }
}

function Get-InstallerFileInfo {
    param([string]$FilesPath)
    
    $installerInfo = @{
        FileName = $null
        Type = $null
    }
    
    # Check for EXE files first
    $exeFiles = Get-ChildItem -Path $FilesPath -Filter "*.exe" -File | Where-Object { $_.Name -notlike "*Setup*" -and $_.Name -notlike "*Invoke-AppDeployToolkit*" }
    if ($exeFiles) {
        $installerInfo.FileName = $exeFiles[0].Name
        $installerInfo.Type = 'exe'
        return $installerInfo
    }
    
    # Check for MSI files
    $msiFiles = Get-ChildItem -Path $FilesPath -Filter "*.msi" -File
    if ($msiFiles) {
        $installerInfo.FileName = $msiFiles[0].Name
        $installerInfo.Type = 'msi'
        return $installerInfo
    }
    
    # Check for MSIX/APPX files
    $msixFiles = Get-ChildItem -Path $FilesPath -Filter "*.msix" -File
    if ($msixFiles) {
        $installerInfo.FileName = $msixFiles[0].Name
        $installerInfo.Type = 'msix'
        return $installerInfo
    }
    
    $appxFiles = Get-ChildItem -Path $FilesPath -Filter "*.appx" -File
    if ($appxFiles) {
        $installerInfo.FileName = $appxFiles[0].Name
        $installerInfo.Type = 'appx'
        return $installerInfo
    }
    
    return $installerInfo
}

function Update-PSADTDeploymentScript {
    param(
        [string]$ScriptPath,
        [hashtable]$InstallerInfo,
        [string]$InstallerFileName,
        [string]$InstallerType
    )
    
    try {
        Write-Host "Updating deployment script: $ScriptPath" -ForegroundColor Green
        
        $scriptContent = Get-Content $ScriptPath -Raw
        
        # Update app variables in $adtSession
        $appVendor = if ($InstallerInfo.Publisher) { "'$($InstallerInfo.Publisher)'" } else { "''" }
        $appName = if ($InstallerInfo.PackageName) { "'$($InstallerInfo.PackageName)'" } else { "''" }
        $appVersion = if ($InstallerInfo.PackageVersion) { "'$($InstallerInfo.PackageVersion)'" } else { "''" }
        $appArch = if ($InstallerInfo.Architecture) { "'$($InstallerInfo.Architecture)'" } else { "''" }
        
        # Replace app variables
        $scriptContent = $scriptContent -replace "AppVendor = ''", "AppVendor = $appVendor"
        $scriptContent = $scriptContent -replace "AppName = ''", "AppName = $appName"
        $scriptContent = $scriptContent -replace "AppVersion = ''", "AppVersion = $appVersion"
        $scriptContent = $scriptContent -replace "AppArch = ''", "AppArch = $appArch"
        
        # For EXE installers, add installation logic
        if ($InstallerType -eq 'exe') {
            $installLogic = @"

    ## Install EXE Application
    `$installerPath = Join-Path `$adtSession.DirFiles '$InstallerFileName'
    if (Test-Path `$installerPath) {
        Write-ADTLogEntry -Message "Installing $($InstallerInfo.PackageName) from: `$installerPath" -Severity 1
        
        # Silent installation
        `$installArgs = '$($InstallerInfo.SilentArgs)'
        Start-ADTProcess -FilePath `$installerPath -ArgumentList `$installArgs -WaitForMsiExec
        
        Write-ADTLogEntry -Message "$($InstallerInfo.PackageName) installation completed" -Severity 1
    } else {
        Write-ADTLogEntry -Message "Installer file not found: `$installerPath" -Severity 3
        throw "Installer file not found"
    }
"@

            # Replace installation section - handle both new and existing installations
            if ($scriptContent -match [regex]::Escape('## <Perform Installation tasks here>')) {
                $scriptContent = $scriptContent -replace [regex]::Escape('## <Perform Installation tasks here>'), $installLogic
            }
            elseif ($scriptContent -match '## Install EXE Application') {
                # Find and replace the entire EXE installation block
                $pattern = '## Install EXE Application[\s\S]*?(?=\n\s*##\s*=|\z)'
                $scriptContent = $scriptContent -replace $pattern, ($installLogic + "`n")
            }
        }
        
        # Write updated content back to file
        Set-Content -Path $ScriptPath -Value $scriptContent -Encoding UTF8
        Write-Host "Successfully updated deployment script!" -ForegroundColor Green
        
    }
    catch {
        Write-Error "Failed to update deployment script: $($_.Exception.Message)"
    }
}

# Main execution
try {
    Write-Host "PSADT EXE Installer Configuration Tool" -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan
    
    # Validate paths
    if (-not (Test-Path $ProjectPath)) {
        throw "Project path does not exist: $ProjectPath"
    }
    
    if (-not (Test-Path $FilesPath)) {
        throw "Files path does not exist: $FilesPath"
    }
    
    if (-not (Test-Path $DeploymentScriptPath)) {
        throw "Deployment script not found: $DeploymentScriptPath"
    }
    
    Write-Host "Scanning Files folder for YAML configurations..." -ForegroundColor Yellow
    
    # Find YAML files
    $yamlFiles = Get-ChildItem -Path $FilesPath -Filter "*.yaml" -File
    
    if ($yamlFiles.Count -eq 0) {
        Write-Host "No YAML files found in Files folder" -ForegroundColor Yellow
        return
    }
    
    foreach ($yamlFile in $yamlFiles) {
        Write-Host "`nProcessing: $($yamlFile.Name)" -ForegroundColor Cyan
        
        # Parse YAML
        $installerInfo = Get-YAMLInstallerInfo -YAMLFilePath $yamlFile.FullName
        
        if (-not $installerInfo) {
            Write-Warning "Failed to parse YAML file: $($yamlFile.Name)"
            continue
        }
        
        # Detect installer type by examining actual files in Files folder
        $fileInfo = Get-InstallerFileInfo -FilesPath $FilesPath
        
        if (-not $fileInfo.FileName) {
            Write-Warning "No installer files (EXE/MSI/MSIX) found in Files folder"
            continue
        }
        
        Write-Host "Detected Installer Type: $($fileInfo.Type.ToUpper())" -ForegroundColor White
        Write-Host "Installer File: $($fileInfo.FileName)" -ForegroundColor White
        Write-Host "Package Name: $($installerInfo.PackageName)" -ForegroundColor White
        Write-Host "Version: $($installerInfo.PackageVersion)" -ForegroundColor White
        Write-Host "Publisher: $($installerInfo.Publisher)" -ForegroundColor White
        Write-Host "Silent Args: $($installerInfo.SilentArgs)" -ForegroundColor White
        
        # Process based on detected file type
        if ($fileInfo.Type -eq 'exe') {
            Write-Host "EXE installer detected - updating deployment script..." -ForegroundColor Yellow
            
            # Update the deployment script
            Update-PSADTDeploymentScript -ScriptPath $DeploymentScriptPath -InstallerInfo $installerInfo -InstallerFileName $fileInfo.FileName -InstallerType $fileInfo.Type
        }
        elseif ($fileInfo.Type -eq 'msi') {
            Write-Host "MSI installer detected - Zero-Config MSI will handle installation automatically" -ForegroundColor Green
        }
        elseif ($fileInfo.Type -in @('msix', 'appx')) {
            Write-Host "$($fileInfo.Type.ToUpper()) installer detected - manual configuration may be required" -ForegroundColor Yellow
        }
        else {
            Write-Host "Unknown installer type detected" -ForegroundColor Yellow
        }
    }
    
    Write-Host "`nConfiguration completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
}