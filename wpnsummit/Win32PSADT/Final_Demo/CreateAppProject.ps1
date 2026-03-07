#Requires -Version 5.1

<#
.SYNOPSIS
    Complete PSADT Project Creation Solution
.DESCRIPTION
    Comprehensive script that:
    1. Searches for Winget applications with architecture selection
    2. Downloads applications automatically  
    3. Creates PSADT V4 projects with intelligent installer configuration
    4. Generates Intune requirement scripts for deployment prerequisites
    5. Creates targeted installation documentation using Windows Sandbox
    
    This script provides end-to-end automation from app discovery to deployment-ready PSADT projects.
.PARAMETER SearchTerm
    Term to search for in Winget repository
.PARAMETER Id
    Winget package ID to use directly (skips search). Example: 'Git.Git'
.PARAMETER BasePath
    Base path where downloads, PSADT projects, and documentation will be created
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SearchTerm,

    [Parameter(Mandatory = $false)]
    [string]$Id,
    
    [Parameter(Mandatory = $false)]
    [string]$BasePath = "C:\Github\Speaker_Sessions\wpnsummit\Win32PSADT\Final_Demo"
)

function Test-WingetInstalled {
    try {
        $null = Get-Command winget -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Search-WingetApps {
    param([string]$SearchTerm)
    
    Write-Host "Searching for apps matching: $SearchTerm" -ForegroundColor Yellow
    
    # Run winget search and capture output
    $searchResults = winget search $SearchTerm --accept-source-agreements | Out-String
    
    # Parse the results (skip header lines)
    $lines = $searchResults -split "`n" | Where-Object { $_.Trim() -ne "" }
    $apps = @()
    
    # Find the separator line (dashes) to know where data starts; handle varying winget formats
    $dataStartIndex = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^[-\s]+$" -and $lines[$i] -match "-{2,}") {
            $dataStartIndex = $i + 1
            break
        }
    }
    
    # Parse each app line
    for ($i = $dataStartIndex; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if ($line -and
            $line -notmatch "^\d+\s+matches?\s+found" -and
            $line -notmatch "^More\s+than" -and
            $line -notmatch "^Name\s+Id\s+" ) {   # skip header row if separator detection missed it
            # Split by multiple spaces to separate columns
            $parts = $line -split '\s{2,}'
            if ($parts.Count -ge 3) {
                $apps += [PSCustomObject]@{
                    Name    = $parts[0].Trim()
                    Id      = $parts[1].Trim()
                    Version = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "" }
                    Source  = if ($parts.Count -gt 3) { $parts[$parts.Count - 1].Trim() } else { "winget" }
                }
            }
        }
    }
    
    return $apps
}

function Get-WingetAppDetails {
    param([string]$AppId)
    
    Write-Host "Getting details for: $AppId" -ForegroundColor Yellow
    
    try {
        # Get app information including available architectures
        $appInfo = winget show $AppId --accept-source-agreements | Out-String
        
        # Parse architectures from the output
        $architectures = @()
        $lines = $appInfo -split "`n"
        
        foreach ($line in $lines) {
            if ($line -match "Architecture:\s*(.+)") {
                $archList = $matches[1] -split ',' | ForEach-Object { $_.Trim() }
                $architectures += $archList
            }
        }
        
        # Remove duplicates and filter out empty values
        $architectures = $architectures | Where-Object { $_ } | Sort-Object -Unique
        
        # If no specific architectures found, try to get them from the installer info
        if ($architectures.Count -eq 0) {
            # Look for installer information that might contain architecture details
            foreach ($line in $lines) {
                if ($line -match "Installer\s+Type:\s*(.+)" -or $line -match "Package\s+Identifier:" -or $line -match "Platform:") {
                    # Check if ARM64 is mentioned anywhere in the app info
                    if ($appInfo -match "arm64|aarch64" -or $appInfo -match "ARM64") {
                        $architectures += "arm64"
                    }
                    if ($appInfo -match "x64|amd64" -or $appInfo -match "x86_64") {
                        $architectures += "x64"
                    }
                    if ($appInfo -match "x86|i386" -and $appInfo -notmatch "x86_64") {
                        $architectures += "x86"
                    }
                }
            }
            
            # Remove duplicates again after parsing
            $architectures = $architectures | Where-Object { $_ } | Sort-Object -Unique
            
            # Final fallback - provide all common architectures as options
            if ($architectures.Count -eq 0) {
                $architectures = @("x64", "x86", "arm64")
                Write-Host "No specific architectures detected. Showing common options." -ForegroundColor Yellow
            }
        }
        
        return $architectures
    }
    catch {
        Write-Warning "Could not get app details for $AppId : $($_.Exception.Message)"
        return @("x64", "x86")  # Default architectures
    }
}

function Select-Architecture {
    param(
        [array]$Architectures,
        [string]$AppName
    )
    
    # Ensure we always have ARM64 as an option if it's not already detected
    $allArchOptions = @("x64", "x86", "arm64")
    $finalArchs = @()
    
    # Add detected architectures first
    foreach ($arch in $Architectures) {
        if ($arch -notin $finalArchs) {
            $finalArchs += $arch
        }
    }
    
    # Add any missing common architectures
    foreach ($common in $allArchOptions) {
        if ($common -notin $finalArchs) {
            $finalArchs += $common
        }
    }
    
    Write-Host "`nArchitecture options for $AppName :" -ForegroundColor Cyan
    Write-Host "(* indicates detected as available)" -ForegroundColor Gray
    
    for ($i = 0; $i -lt $finalArchs.Count; $i++) {
        $marker = if ($finalArchs[$i] -in $Architectures) { "*" } else { " " }
        Write-Host "  $($i + 1). $($finalArchs[$i]) $marker" -ForegroundColor White
    }
    Write-Host "  $($finalArchs.Count + 1). All detected architectures" -ForegroundColor White
    
    do {
        $selection = Read-Host "`nSelect architecture (1-$($finalArchs.Count + 1))"
        
        if ([int]$selection -ge 1 -and [int]$selection -le $finalArchs.Count) {
            return $finalArchs[$selection - 1]
        }
        elseif ([int]$selection -eq ($finalArchs.Count + 1)) {
            return "all"
        }
        
        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
    } while ($true)
}

function Download-WingetApp {
    param(
        [string]$AppId,
        [string]$AppName,
        [string]$DownloadPath,
        [string]$Architecture = $null
    )
    
    Write-Host "Downloading $AppName ($AppId) [$Architecture]..." -ForegroundColor Green
    
    try {
        # Build download command
        $downloadCmd = "winget download --id `"$AppId`" --download-directory `"$DownloadPath`" --accept-source-agreements --accept-package-agreements"
        
        # Add architecture if specified
        if ($Architecture -and $Architecture -ne "all") {
            $downloadCmd += " --architecture `"$Architecture`""
        }
        
        Write-Host "Executing: $downloadCmd" -ForegroundColor Gray
        Invoke-Expression $downloadCmd
        
        Write-Host "Successfully downloaded to: $DownloadPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to download $AppName : $($_.Exception.Message)"
        return $false
    }
}

function Create-PSADTProject {
    param(
        [string]$ProjectName,
        [string]$ProjectPath
    )
    
    $ProjectFullPath = Join-Path $ProjectPath $ProjectName
    
    try {
        Write-Host "`nCreating PSADT V4 project: $ProjectName" -ForegroundColor Green
        
        # Check if PSADT module is installed
        Write-Host "Checking for PSAppDeployToolkit module..." -ForegroundColor Yellow
        $PSADTModule = Get-Module -Name PSAppDeployToolkit -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $PSADTModule) {
            Write-Host "PSAppDeployToolkit module not found. Installing..." -ForegroundColor Yellow
            Install-Module -Name PSAppDeployToolkit -Scope CurrentUser -Force -AllowClobber
            $PSADTModule = Get-Module -Name PSAppDeployToolkit -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        }
        else {
            # Check PSGallery for a newer version
            Write-Host "Installed version: $($PSADTModule.Version) — checking PSGallery for updates..." -ForegroundColor Yellow
            try {
                $galleryModule = Find-Module -Name PSAppDeployToolkit -Repository PSGallery -ErrorAction Stop
                $galleryVersion = [System.Version]$galleryModule.Version
                $installedVersion = [System.Version]$PSADTModule.Version

                if ($galleryVersion -gt $installedVersion) {
                    Write-Host "Update available: $galleryVersion (installed: $installedVersion)" -ForegroundColor Cyan
                    $doUpdate = Read-Host "Update PSAppDeployToolkit to $galleryVersion now? (y/n)"
                    if ($doUpdate -in @('y', 'Y')) {
                        Write-Host "Updating PSAppDeployToolkit to $galleryVersion..." -ForegroundColor Yellow
                        Update-Module -Name PSAppDeployToolkit -Force
                        Write-Host "✓ Updated to version: $galleryVersion" -ForegroundColor Green
                        Write-Host ""
                        Write-Host "The module assembly has changed — restarting script in a new PowerShell session..." -ForegroundColor Cyan
                        # Re-launch with the same arguments so the new assembly loads cleanly
                        $scriptArgs = @('-NoExit', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
                        foreach ($key in $PSBoundParameters.Keys) {
                            $val = $PSBoundParameters[$key]
                            if ($val -is [switch]) {
                                if ($val.IsPresent) { $scriptArgs += "-$key" }
                            } else {
                                $scriptArgs += "-$key", "`"$val`""
                            }
                        }
                        Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $scriptArgs
                        exit
                    }
                    else {
                        Write-Host "Skipping update — continuing with installed version $installedVersion" -ForegroundColor Gray
                    }
                }
                else {
                    Write-Host "✓ PSAppDeployToolkit is up to date ($($PSADTModule.Version))" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "Could not reach PSGallery to check for updates: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
        
        Write-Host "Using PSAppDeployToolkit version: $($PSADTModule.Version)" -ForegroundColor Green
        
        # Import the module
        Import-Module -Name PSAppDeployToolkit -Force
        
        # Create project directory structure if it doesn't exist
        if (!(Test-Path $ProjectPath)) {
            New-Item -Path $ProjectPath -ItemType Directory -Force | Out-Null
        }
        
        if (Test-Path $ProjectFullPath) {
            Write-Warning "Project directory already exists: $ProjectFullPath"
            $overwrite = Read-Host "Do you want to overwrite it? (y/n)"
            if ($overwrite -notin @('y', 'Y')) {
                return $false
            }
            Remove-Item -Path $ProjectFullPath -Recurse -Force
        }
        
        # Create new PSADT template using the module
        Write-Host "Creating PSADT V4 template..." -ForegroundColor Yellow
        New-ADTTemplate -Destination $ProjectPath -Name $ProjectName
        
        Write-Host "PSADT V4 project created successfully!" -ForegroundColor Green
        Write-Host "Project location: $ProjectFullPath" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Error "Failed to create PSADT project: $($_.Exception.Message)"
        return $false
    }
}

function Sanitize-ProjectName {
    param([string]$Name)
    
    # Remove or replace invalid characters for folder names
    $sanitized = $Name -replace '[<>:"/\\|?*]', '_'
    $sanitized = $sanitized -replace '\s+', '_'  # Replace spaces with underscores
    $sanitized = $sanitized -replace '_+', '_'   # Replace multiple underscores with single
    $sanitized = $sanitized.Trim('_')            # Remove leading/trailing underscores
    
    return $sanitized
}

function Get-InstallerFileInfo {
    param([string]$FilesPath)
    
    $installerInfo = @{
        FileName = $null
        Type = $null
        FullPath = $null
    }
    
    # Check for MSI files first (Zero-Config MSI priority)
    $msiFiles = Get-ChildItem -Path $FilesPath -Filter "*.msi" -File
    if ($msiFiles) {
        $installerInfo.FileName = $msiFiles[0].Name
        $installerInfo.Type = 'msi'
        $installerInfo.FullPath = $msiFiles[0].FullName
        return $installerInfo
    }
    
    # Check for EXE files
    $exeFiles = Get-ChildItem -Path $FilesPath -Filter "*.exe" -File | Where-Object { 
        $_.Name -notlike "*Setup*" -and 
        $_.Name -notlike "*Invoke-AppDeployToolkit*" -and
        $_.Name -notlike "*ServiceUI*"
    }
    if ($exeFiles) {
        $installerInfo.FileName = $exeFiles[0].Name
        $installerInfo.Type = 'exe'
        $installerInfo.FullPath = $exeFiles[0].FullName
        return $installerInfo
    }
    
    # Check for MSIX/APPX files
    $msixFiles = Get-ChildItem -Path $FilesPath -Filter "*.msix" -File
    if ($msixFiles) {
        $installerInfo.FileName = $msixFiles[0].Name
        $installerInfo.Type = 'msix'
        $installerInfo.FullPath = $msixFiles[0].FullName
        return $installerInfo
    }
    
    $appxFiles = Get-ChildItem -Path $FilesPath -Filter "*.appx" -File
    if ($appxFiles) {
        $installerInfo.FileName = $appxFiles[0].Name
        $installerInfo.Type = 'appx'
        $installerInfo.FullPath = $appxFiles[0].FullName
        return $installerInfo
    }
    
    return $installerInfo
}

function Get-YAMLInstallerInfo {
    param([string]$FilesPath)
    
    $yamlFiles = Get-ChildItem -Path $FilesPath -Filter "*.yaml" -File
    if ($yamlFiles.Count -eq 0) {
        return $null
    }
    
    try {
        $yamlContent = Get-Content $yamlFiles[0].FullName -Raw
        $installerInfo = @{
            PackageName = $null
            Publisher = $null
            PackageVersion = $null
            Architecture = $null
            SilentArgs = $null
            ProductCode = $null
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
        if ($yamlContent -match 'Architecture:\s*(.+)') {
            $installerInfo.Architecture = $matches[1].Trim()
        }
        if ($yamlContent -match 'ProductCode:\s*(.+)') {
            $installerInfo.ProductCode = $matches[1].Trim()
        }
        
        # Parse installer switches (Silent: may appear after other keys under InstallerSwitches:)
        if ($yamlContent -match '(?s)InstallerSwitches:.*?\n\s+Silent:\s*([^\n]+)') {
            $installerInfo.SilentArgs = $matches[1].Trim()
        }
        
        return $installerInfo
    }
    catch {
        Write-Warning "Failed to parse YAML file: $($_.Exception.Message)"
        return $null
    }
}

function Configure-PSADTForInstaller {
    param(
        [string]$ProjectPath,
        [PSCustomObject]$AppInfo,
        [string]$Architecture
    )
    
    try {
        $filesPath = Join-Path $ProjectPath "Files"
        $scriptPath = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
        
        if (-not (Test-Path $scriptPath)) {
            Write-Warning "PSADT script not found: $scriptPath"
            return $false
        }
        
        # Detect installer type
        $fileInfo = Get-InstallerFileInfo -FilesPath $filesPath
        if (-not $fileInfo.FileName) {
            Write-Warning "No installer files detected in Files folder"
            return $false
        }
        
        Write-Host "Detected installer: $($fileInfo.FileName) ($($fileInfo.Type.ToUpper()))" -ForegroundColor Green
        
        # Get YAML info if available
        $yamlInfo = Get-YAMLInstallerInfo -FilesPath $filesPath
        
        # Read current script content
        $scriptContent = Get-Content $scriptPath -Raw
        
        # Prepare app variables
        $appVendor = if ($yamlInfo.Publisher) { "'$($yamlInfo.Publisher)'" } else { "''" }
        $appVersion = if ($yamlInfo.PackageVersion -or $AppInfo.Version) { 
            "'$(if ($yamlInfo.PackageVersion) { $yamlInfo.PackageVersion } else { $AppInfo.Version })'" 
        } else { "''" }
        $appArch = "'$Architecture'"
        
        # Configure based on installer type
        if ($fileInfo.Type -eq 'msi') {
            Write-Host "Configuring for MSI installer (Zero-Config MSI)" -ForegroundColor Yellow
            
            # For MSI: Keep AppName empty to enable Zero-Config MSI
            $appName = "''"
            
            # Update app variables
            $scriptContent = $scriptContent -replace "AppVendor = ''", "AppVendor = $appVendor"
            $scriptContent = $scriptContent -replace "AppVersion = ''", "AppVersion = $appVersion"
            $scriptContent = $scriptContent -replace "AppArch = ''", "AppArch = $appArch"
            
            Write-Host "✓ MSI Zero-Config enabled (AppName left empty)" -ForegroundColor Green
        }
        elseif ($fileInfo.Type -eq 'exe') {
            Write-Host "Configuring for EXE installer" -ForegroundColor Yellow
            
            # For EXE: Set AppName to disable Zero-Config and add install logic
            $appName = if ($yamlInfo.PackageName) { "'$($yamlInfo.PackageName)'" } else { "'$($AppInfo.Name)'" }
            
            # Update app variables  
            $scriptContent = $scriptContent -replace "AppVendor = ''", "AppVendor = $appVendor"
            $scriptContent = $scriptContent -replace "AppName = ''", "AppName = $appName"
            $scriptContent = $scriptContent -replace "AppVersion = ''", "AppVersion = $appVersion"
            $scriptContent = $scriptContent -replace "AppArch = ''", "AppArch = $appArch"
            
            # Add EXE installation logic using proven approach from Update-PSADTForEXE.ps1
            $silentArgs = if ($yamlInfo.SilentArgs) { $yamlInfo.SilentArgs } else { "/S" }
            $packageName = if ($yamlInfo.PackageName) { $yamlInfo.PackageName } else { $AppInfo.Name }
            
            $installLogic = @"

    ## Install EXE Application
    `$installerPath = Join-Path `$adtSession.DirFiles '$($fileInfo.FileName)'
    if (Test-Path `$installerPath) {
        Write-ADTLogEntry -Message "Installing $packageName from: `$installerPath" -Severity 1
        
        # Silent installation
        `$installArgs = '$silentArgs'
        Start-ADTProcess -FilePath `$installerPath -ArgumentList `$installArgs
        
        Write-ADTLogEntry -Message "$packageName installation completed" -Severity 1
    } else {
        Write-ADTLogEntry -Message "Installer file not found: `$installerPath" -Severity 3
        throw "Installer file not found"
    }
"@
            
            # Replace installation section - handle both new and existing installations  
            if ($scriptContent -match [regex]::Escape('## <Perform Installation tasks here>')) {
                $scriptContent = $scriptContent -replace [regex]::Escape('## <Perform Installation tasks here>'), $installLogic
                Write-Host "✓ EXE installation logic added" -ForegroundColor Green
            }
            elseif ($scriptContent -match '## Install EXE Application') {
                # Find and replace the entire EXE installation block
                $pattern = '## Install EXE Application[\s\S]*?(?=\n\s*##\s*=|\z)'
                $scriptContent = $scriptContent -replace $pattern, ($installLogic + "`n")
                Write-Host "✓ EXE installation logic updated" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Configuring for $($fileInfo.Type.ToUpper()) installer" -ForegroundColor Yellow
            
            # For other types: Set AppName and basic variables
            $appName = if ($yamlInfo.PackageName) { "'$($yamlInfo.PackageName)'" } else { "'$($AppInfo.Name)'" }
            
            $scriptContent = $scriptContent -replace "AppVendor = ''", "AppVendor = $appVendor"
            $scriptContent = $scriptContent -replace "AppName = ''", "AppName = $appName"
            $scriptContent = $scriptContent -replace "AppVersion = ''", "AppVersion = $appVersion"
            $scriptContent = $scriptContent -replace "AppArch = ''", "AppArch = $appArch"
            
            Write-Host "✓ Basic app variables configured" -ForegroundColor Green
        }
        
        # Update script date
        $currentDate = Get-Date -Format 'yyyy-MM-dd'
        $scriptContent = $scriptContent -replace "AppScriptDate = '2025-09-23'", "AppScriptDate = '$currentDate'"
        
        # Write updated content back to file
        Set-Content -Path $scriptPath -Value $scriptContent -Encoding UTF8

        # Apply org template branding and dialog settings
        if ($script:OrgTemplate) {
            Write-Host 'Applying org template...' -ForegroundColor Cyan
            Apply-OrgTemplate -ProjectPath $ProjectPath -Template $script:OrgTemplate | Out-Null
        }

        Write-Host "✓ PSADT script configured successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to configure PSADT: $($_.Exception.Message)"
        return $false
    }
}

function New-TargetedDocumentation {
    param(
        [string]$ProjectPath,
        [string]$ProjectName,
        [object]$AppInfo
    )
    
    try {
        Write-Host "Creating targeted documentation configuration..." -ForegroundColor Yellow
        
        # Create unique timestamp for this documentation session
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Define paths
        $supportFilesFolder = Join-Path $ProjectPath "SupportFiles"
        $sandboxConfigFile = Join-Path $ProjectPath "${ProjectName}_TargetedDocumentation.wsb"
        
        # Ensure SupportFiles folder exists
        if (-not (Test-Path $supportFilesFolder)) {
            New-Item -ItemType Directory -Path $supportFilesFolder -Force | Out-Null
        }
        
        # Create the documentation script content
        $documentationScript = @'
# Targeted PSADT Installation Documentation Script
# Auto-generated by CreateAppProject.ps1

$projectName = '__PROJECTNAME__'
$timestamp = '__TIMESTAMP__'
$outputPath = 'C:\PSADT\Documentation'
$logPath = 'C:\PSADT\SupportFiles\Targeted_Documentation_Log___TIMESTAMP__.txt'

Write-Host "======================================"
Write-Host "PSADT Installation Documentation (Targeted)"
Write-Host "Project: $projectName"
Write-Host "Timestamp: $timestamp"
Write-Host "======================================"

Write-Host "`nWaiting for Windows Sandbox to fully initialize..." -ForegroundColor Yellow
Write-Host "This ensures all system services are ready before documentation begins." -ForegroundColor Gray
Write-Host "Please wait 60 seconds..." -ForegroundColor Cyan

# Wait 60 seconds for sandbox initialization
for ($i = 60; $i -gt 0; $i--) {
    Write-Progress -Activity "Initializing Windows Sandbox" -Status "Waiting for system to stabilize..." -SecondsRemaining $i -PercentComplete ((60 - $i) / 60 * 100)
    Start-Sleep -Seconds 1
}
Write-Progress -Activity "Initializing Windows Sandbox" -Completed
Write-Host "Sandbox initialization complete!" -ForegroundColor Green

# Create documentation output folder
New-Item -Path $outputPath -ItemType Directory -Force | Out-Null

# Initialize logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
}

# Start logging
Write-Log "========================================" "INFO"
Write-Log "PSADT Installation Documentation Started (Targeted)" "INFO"
Write-Log "Project: $projectName" "INFO"
Write-Log "Timestamp: $timestamp" "INFO"
Write-Log "Output Path: $outputPath" "INFO"
Write-Log "Log Path: $logPath" "INFO"
Write-Log "Windows Sandbox initialization wait started (60 seconds)" "INFO"
Write-Log "Waiting for system services to stabilize before baseline capture" "INFO"
Write-Log "Windows Sandbox initialization completed successfully" "INFO"
Write-Log "========================================" "INFO"

# Function to export data with error handling and proper formatting
function Export-DataSafely {
    param([string]$FilePath, [object]$Data, [string]$Description)
    try {
        $itemCount = if ($Data -is [array]) { $Data.Count } else { 1 }
        Write-Log "Attempting to export $Description ($itemCount items) to: $FilePath" "INFO"
        
        # Use Out-String with width to prevent truncation, then save to file
        $Data | Out-String -Width 4096 | Set-Content -Path $FilePath -Encoding UTF8 -Force
        
        Write-Host "Exported: $Description" -ForegroundColor Green
        Write-Log "Successfully exported $Description ($itemCount items)" "SUCCESS"
    } catch {
        Write-Host "Failed to export $Description - $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Failed to export $Description - $($_.Exception.Message)" "ERROR"
    }
}

# Function to scan targeted directories
function Get-TargetedDirectorySnapshot {
    param([string]$Description, [string[]]$Paths, [int]$Depth = 3)
    
    Write-Host "Scanning $Description..." -ForegroundColor Gray
    Write-Log "Beginning $Description scan with depth $Depth" "INFO"
    
    $results = @()
    
    foreach ($path in $Paths) {
        if (Test-Path $path) {
            try {
                Write-Log "Scanning path: $path" "INFO"
                $items = Get-ChildItem -Path $path -Recurse -Depth $Depth -ErrorAction SilentlyContinue | 
                         Select-Object FullName, Name, Length, CreationTime, LastWriteTime, Attributes
                $results += $items
                Write-Log "Found $($items.Count) items in $path" "INFO"
            } catch {
                Write-Log "Error scanning $path - $($_.Exception.Message)" "WARNING"
            }
        } else {
            Write-Log "Path not found: $path" "INFO"
        }
    }
    
    Write-Log "Total items found in $Description - $($results.Count)" "SUCCESS"
    return $results
}

# Function to scan targeted registry locations
function Get-TargetedRegistrySnapshot {
    param([string]$Description, [string[]]$RegistryPaths)
    
    Write-Host "Scanning $Description..." -ForegroundColor Gray
    Write-Log "Beginning $Description scan" "INFO"
    
    $results = @()
    
    foreach ($regPath in $RegistryPaths) {
        try {
            Write-Log "Scanning registry path: $regPath" "INFO"
            
            if (Test-Path $regPath) {
                # Get all subkeys and their properties
                $keys = Get-ChildItem -Path $regPath -Recurse -ErrorAction SilentlyContinue
                
                foreach ($key in $keys) {
                    try {
                        $properties = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                        if ($properties) {
                            # Create detailed registry entry with full paths and values
                            $propList = $properties.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" }
                            $propString = ""
                            foreach ($prop in $propList) {
                                $value = if ($prop.Value -is [array]) { $prop.Value -join ", " } else { $prop.Value }
                                $propString += "`n    $($prop.Name) = $value"
                            }
                            
                            $results += [PSCustomObject]@{
                                FullPath = $key.PSPath -replace "Microsoft.PowerShell.Core\\Registry::", ""
                                KeyName = $key.PSChildName
                                Properties = $propString
                                Summary = "Key: $($key.PSChildName) | Properties: $($propList.Count)"
                            }
                        }
                    } catch {
                        Write-Log "Error reading properties from $($key.PSPath) - $($_.Exception.Message)" "WARNING"
                    }
                }
                
                Write-Log "Found $($keys.Count) registry keys in $regPath" "INFO"
            } else {
                Write-Log "Registry path not found: $regPath" "INFO"
            }
        } catch {
            Write-Log "Error scanning registry path $regPath - $($_.Exception.Message)" "WARNING"
        }
    }
    
    Write-Log "Total registry entries found in $Description - $($results.Count)" "SUCCESS"
    return $results
}

Write-Host "`nCapturing PRE-INSTALLATION baseline..." -ForegroundColor Yellow
Write-Log "Starting PRE-INSTALLATION baseline capture" "INFO"

# Define targeted directories to monitor
$targetDirectories = @(
    "${env:ProgramFiles}",
    "${env:ProgramFiles(x86)}",
    "${env:ProgramData}",
    "${env:LOCALAPPDATA}",
    "${env:APPDATA}"
)

# Define targeted registry locations
$targetRegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths", 
    "HKLM:\SOFTWARE\Classes\Applications",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Capture baseline file system state (no export - keep in memory only)
$preFiles = Get-TargetedDirectorySnapshot -Description "Pre-installation files" -Paths $targetDirectories -Depth 3
Write-Log "Pre-installation files captured: $($preFiles.Count) files" "INFO"

# Capture baseline registry state (no export - keep in memory only)
$preRegistry = Get-TargetedRegistrySnapshot -Description "Pre-installation registry" -RegistryPaths $targetRegistryPaths
Write-Log "Pre-installation registry captured: $($preRegistry.Count) keys" "INFO"

# Capture baseline services (keep in memory only)
Write-Host "Capturing baseline services..." -ForegroundColor Gray
Write-Log "Beginning baseline services scan" "INFO"
$preServices = Get-Service | Select-Object Name, DisplayName, Status, StartType
Write-Log "Pre-installation services captured: $($preServices.Count) services" "INFO"

# Capture baseline programs (keep in memory only)
Write-Host "Capturing baseline programs..." -ForegroundColor Gray
Write-Log "Beginning baseline programs scan" "INFO"
$prePrograms = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Select-Object Name, Version, Vendor, InstallDate
Write-Log "Pre-installation programs captured: $($prePrograms.Count) programs" "INFO"

Write-Host "`nStarting PSADT installation..." -ForegroundColor Yellow
Write-Log "Starting PSADT installation phase" "INFO"
Write-Host "===============================================" -ForegroundColor Cyan

# Change to PSADT directory
Set-Location C:\PSADT
Write-Log "Changed directory to C:\PSADT" "INFO"

Write-Host "`nThe installation will now begin." -ForegroundColor Yellow
Write-Host "Press any key when ready to start the installation..." -ForegroundColor Cyan
Read-Host

# Run the PSADT installation
Write-Host "Running: .\Invoke-AppDeployToolkit.ps1" -ForegroundColor White
Write-Log "Launching PSADT installation: .\Invoke-AppDeployToolkit.ps1" "INFO"

# Start installation and wait for completion
Write-Host "Starting installation..." -ForegroundColor Gray
Write-Host "Waiting for installation to complete (this may take several minutes)..." -ForegroundColor Yellow

$installProcess = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy", "Bypass", "-File", ".\Invoke-AppDeployToolkit.ps1" -PassThru -Wait

Write-Host "Installation process completed with exit code: $($installProcess.ExitCode)" -ForegroundColor Green
Write-Log "PSADT installation process completed with exit code: $($installProcess.ExitCode)" "INFO"

Write-Host "`nWaiting for installation to fully finalize..." -ForegroundColor Yellow
Write-Host "Allowing time for background processes, registry updates, and file operations to complete." -ForegroundColor Gray
Write-Host "Please wait 30 seconds..." -ForegroundColor Cyan

# Wait 30 seconds for installation finalization - registry changes can be delayed
for ($i = 30; $i -gt 0; $i--) {
    Write-Progress -Activity "Finalizing Installation" -Status "Waiting for background processes and registry updates to complete..." -SecondsRemaining $i -PercentComplete ((30 - $i) / 30 * 100)
    Start-Sleep -Seconds 1
}
Write-Progress -Activity "Finalizing Installation" -Completed
Write-Host "Installation finalization complete!" -ForegroundColor Green

Write-Log "Installation finalization wait completed (30 seconds)" "INFO"

Write-Host "`nCapturing POST-INSTALLATION state..." -ForegroundColor Yellow
Write-Log "Starting POST-INSTALLATION state capture" "INFO"

# Capture post-installation file system state
$postFiles = Get-TargetedDirectorySnapshot -Description "Post-installation files" -Paths $targetDirectories -Depth 3
Write-Log "Post-installation files captured: $($postFiles.Count) files" "INFO"

# Capture post-installation registry state
$postRegistry = Get-TargetedRegistrySnapshot -Description "Post-installation registry" -RegistryPaths $targetRegistryPaths  
Write-Log "Post-installation registry captured: $($postRegistry.Count) keys" "INFO"

# Capture post-installation services
Write-Host "Scanning post-installation services..." -ForegroundColor Gray
Write-Log "Beginning post-installation services scan" "INFO"
$postServices = Get-Service | Select-Object Name, DisplayName, Status, StartType
Write-Log "Post-installation services captured: $($postServices.Count) services" "INFO"

# Capture post-installation programs
Write-Host "Scanning post-installation programs..." -ForegroundColor Gray
Write-Log "Beginning post-installation programs scan" "INFO"
$postPrograms = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Select-Object Name, Version, Vendor, InstallDate
Write-Log "Post-installation programs captured: $($postPrograms.Count) programs" "INFO"

Write-Host "`nAnalyzing changes..." -ForegroundColor Yellow
Write-Log "Starting change analysis" "INFO"

# Compare file system changes
try {
    Write-Log "Comparing file system changes" "INFO"
    $newFiles = Compare-Object -ReferenceObject $preFiles.FullName -DifferenceObject $postFiles.FullName | 
                Where-Object { $_.SideIndicator -eq '=>' } | Select-Object @{Name='NewFile';Expression={$_.InputObject}}
    Write-Log "Found $($newFiles.Count) new files" "INFO"
    
    $modifiedFiles = Compare-Object -ReferenceObject $preFiles -DifferenceObject $postFiles -Property FullName, LastWriteTime | 
                     Where-Object { $_.SideIndicator -eq '=>' -and $_.FullName -in $preFiles.FullName } | 
                     Select-Object @{Name='ModifiedFile';Expression={$_.FullName}}, LastWriteTime
    Write-Log "Found $($modifiedFiles.Count) modified files" "INFO"
} catch {
    Write-Host "Warning: Could not compare file changes" -ForegroundColor Yellow
    Write-Log "Warning: Could not compare file changes: $($_.Exception.Message)" "WARNING"
}

# Compare registry changes with retry logic
try {
    Write-Log "Comparing registry changes" "INFO"
    Write-Host "Registry comparison: PRE entries = $($preRegistry.Count), POST entries = $($postRegistry.Count)" -ForegroundColor Gray
    
    $retryCount = 0
    $maxRetries = 3
    $newRegEntries = @()
    
    do {
        # Compare using FullPath property
        $newRegEntries = Compare-Object -ReferenceObject $preRegistry.FullPath -DifferenceObject $postRegistry.FullPath | 
                         Where-Object { $_.SideIndicator -eq '=>' } | 
                         ForEach-Object { 
                             $newPath = $_.InputObject
                             $fullRegEntry = $postRegistry | Where-Object { $_.FullPath -eq $newPath } | Select-Object -First 1
                             [PSCustomObject]@{
                                 NewRegistryEntry = $newPath
                                 KeyName = $fullRegEntry.KeyName
                                 Properties = $fullRegEntry.Properties
                                 Summary = $fullRegEntry.Summary
                             }
                         }
        
        if ($newRegEntries.Count -eq 0 -and $retryCount -lt $maxRetries) {
            $retryCount++
            Write-Host "No registry changes detected on attempt $retryCount. Waiting 30 seconds for delayed registry writes..." -ForegroundColor Yellow
            Write-Log "Registry scan attempt $retryCount found no changes - waiting 30 seconds for retry" "INFO"
            
            # Wait 30 seconds for delayed registry changes
            for ($i = 30; $i -gt 0; $i--) {
                Write-Progress -Activity "Waiting for Registry Changes" -Status "Registry writes may be delayed - waiting for retry $retryCount/$maxRetries..." -SecondsRemaining $i -PercentComplete ((30 - $i) / 30 * 100)
                Start-Sleep -Seconds 1
            }
            Write-Progress -Activity "Waiting for Registry Changes" -Completed
            
            # Re-capture registry state
            Write-Host "Re-scanning registry state (attempt $($retryCount + 1))..." -ForegroundColor Cyan
            Write-Log "Re-capturing registry state for retry attempt $($retryCount + 1)" "INFO"
            $postRegistry = Get-TargetedRegistrySnapshot -Description "Post-installation registry (retry $($retryCount + 1))" -RegistryPaths $targetRegistryPaths
            Write-Log "Registry re-scan completed: $($postRegistry.Count) keys" "INFO"
        }
    } while ($newRegEntries.Count -eq 0 -and $retryCount -lt $maxRetries)
    
    if ($newRegEntries.Count -gt 0) {
        Write-Host "Found $($newRegEntries.Count) new registry entries" -ForegroundColor Green
        Write-Log "Found $($newRegEntries.Count) new registry entries after $retryCount retries" "INFO"
    } else {
        Write-Host "No registry changes detected after $maxRetries attempts" -ForegroundColor Yellow
        Write-Log "No registry changes detected after $maxRetries retry attempts" "WARNING"
    }
    
} catch {
    Write-Host "Warning: Could not compare registry changes" -ForegroundColor Yellow
    Write-Log "Warning: Could not compare registry changes: $($_.Exception.Message)" "WARNING"
}

# Compare services
try {
    Write-Log "Comparing service changes" "INFO"
    $newServices = Compare-Object -ReferenceObject $preServices.Name -DifferenceObject $postServices.Name | 
                    Where-Object { $_.SideIndicator -eq '=>' } | Select-Object @{Name='NewService';Expression={$_.InputObject}}
    Write-Log "Found $($newServices.Count) new services" "INFO"
} catch {
    Write-Host "Warning: Could not compare service changes" -ForegroundColor Yellow
    Write-Log "Warning: Could not compare service changes: $($_.Exception.Message)" "WARNING"
}

# Compare programs
try {
    Write-Log "Comparing program changes" "INFO"
    if ($prePrograms -and $postPrograms) {
        $newPrograms = Compare-Object -ReferenceObject $prePrograms.Name -DifferenceObject $postPrograms.Name | 
                        Where-Object { $_.SideIndicator -eq '=>' } | Select-Object @{Name='NewProgram';Expression={$_.InputObject}}
        Write-Log "Found $($newPrograms.Count) new programs" "INFO"
    } else {
        Write-Log "Warning: Unable to compare programs - insufficient baseline data" "WARNING"
    }
} catch {
    Write-Host "Warning: Could not compare program changes" -ForegroundColor Yellow
    Write-Log "Warning: Could not compare program changes: $($_.Exception.Message)" "WARNING"
}

# Create JSON output for automation (ONLY output besides log)
Write-Host "`nGenerating JSON documentation for automation..." -ForegroundColor Cyan
Write-Log "Creating JSON documentation for new files and registry keys" "INFO"

# Prepare JSON data structure
$jsonData = @{
    InstallationInfo = @{
        ProjectName = $projectName
        Timestamp = $timestamp
        DocumentationDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    NewFiles = @()
    ModifiedFiles = @()
    NewRegistryKeys = @()
    NewServices = @()
    NewPrograms = @()
}

# Process new files for JSON
if ($newFiles -and $newFiles.Count -gt 0) {
    foreach ($file in $newFiles) {
        $fileInfo = @{
            Path = $file.NewFile
            Size = if (Test-Path $file.NewFile) { (Get-Item $file.NewFile -ErrorAction SilentlyContinue).Length } else { $null }
            CreatedDate = if (Test-Path $file.NewFile) { (Get-Item $file.NewFile -ErrorAction SilentlyContinue).CreationTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            Type = if (Test-Path $file.NewFile) { 
                if ((Get-Item $file.NewFile -ErrorAction SilentlyContinue).PSIsContainer) { "Directory" } else { "File" }
            } else { "Unknown" }
        }
        $jsonData.NewFiles += $fileInfo
    }
}

# Process modified files for JSON
if ($modifiedFiles -and $modifiedFiles.Count -gt 0) {
    foreach ($file in $modifiedFiles) {
        $fileInfo = @{
            Path = $file.ModifiedFile
            Size = if (Test-Path $file.ModifiedFile) { (Get-Item $file.ModifiedFile -ErrorAction SilentlyContinue).Length } else { $null }
            ModifiedDate = if (Test-Path $file.ModifiedFile) { (Get-Item $file.ModifiedFile -ErrorAction SilentlyContinue).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
        }
        $jsonData.ModifiedFiles += $fileInfo
    }
}

# Process new registry entries for JSON
if ($newRegEntries -and $newRegEntries.Count -gt 0) {
    Write-Log "Processing $($newRegEntries.Count) new registry entries for JSON" "INFO"
    foreach ($regEntry in $newRegEntries) {
        # Get the full registry information for this path
        $regPath = $regEntry.NewRegistryEntry
        Write-Log "Processing registry path for JSON: $regPath" "INFO"
        try {
            $regItem = Get-Item -Path "Registry::$regPath" -ErrorAction SilentlyContinue
            if ($regItem) {
                $regInfo = @{
                    Path = $regPath
                    KeyName = $regEntry.KeyName
                    ValueCount = $regItem.ValueCount
                    SubKeyCount = $regItem.SubKeyCount
                    Values = @{}
                    Properties = $regEntry.Properties
                }
                
                # Get all values in this registry key
                try {
                    $regValues = Get-ItemProperty -Path "Registry::$regPath" -ErrorAction SilentlyContinue
                    if ($regValues) {
                        $regValues.PSObject.Properties | ForEach-Object {
                            if ($_.Name -notin @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider')) {
                                $regInfo.Values[$_.Name] = $_.Value
                            }
                        }
                    }
                } catch {
                    # Registry key might not have values or access denied
                }
                
                $jsonData.NewRegistryKeys += $regInfo
            }
        } catch {
            # Fallback for inaccessible registry keys - still include basic info
            Write-Log "Registry key not accessible, using fallback data: $regPath" "WARNING"
            $regInfo = @{
                Path = $regPath
                KeyName = $regEntry.KeyName
                ValueCount = "Unknown"
                SubKeyCount = "Unknown"
                Values = @{}
                Properties = $regEntry.Properties
                Note = "Access denied or key not accessible"
            }
            $jsonData.NewRegistryKeys += $regInfo
        }
    }
} else {
    Write-Log "No new registry entries found for JSON processing" "INFO"
    Write-Host "No new registry entries detected" -ForegroundColor Yellow
}

# Process new services for JSON
if ($newServices -and $newServices.Count -gt 0) {
    foreach ($service in $newServices) {
        $serviceInfo = @{
            Name = $service.NewService
            DisplayName = (Get-Service -Name $service.NewService -ErrorAction SilentlyContinue).DisplayName
            Status = (Get-Service -Name $service.NewService -ErrorAction SilentlyContinue).Status
            StartType = (Get-Service -Name $service.NewService -ErrorAction SilentlyContinue).StartType
        }
        $jsonData.NewServices += $serviceInfo
    }
}

# Process new programs for JSON
if ($newPrograms -and $newPrograms.Count -gt 0) {
    foreach ($program in $newPrograms) {
        $programInfo = @{
            Name = $program.NewProgram
            Path = $null  # Could be enhanced to find installation path
        }
        $jsonData.NewPrograms += $programInfo
    }
}

# Export JSON data
try {
    $jsonOutput = $jsonData | ConvertTo-Json -Depth 10 -Compress:$false
    $jsonFilePath = "$outputPath\InstallationChanges_$timestamp.json"
    Set-Content -Path $jsonFilePath -Value $jsonOutput -Encoding UTF8
    Write-Host "JSON documentation exported: InstallationChanges_$timestamp.json" -ForegroundColor Green
    Write-Log "JSON documentation exported successfully to: $jsonFilePath" "SUCCESS"
    
} catch {
    Write-Host "Warning: Could not export JSON documentation" -ForegroundColor Yellow
    Write-Log "Warning: Could not export JSON documentation: $($_.Exception.Message)" "WARNING"
}

# Copy log file to documentation output folder
try {
    Copy-Item -Path $logPath -Destination "$outputPath\Targeted_Documentation_Log_$timestamp.txt" -Force
    Write-Log "Log file copied to documentation output folder" "SUCCESS"
} catch {
    Write-Log "Failed to copy log file to documentation folder: $($_.Exception.Message)" "WARNING"
}

Write-Host "`n======================================"
Write-Host "Targeted Documentation completed successfully!" -ForegroundColor Green
Write-Log "========================================" "SUCCESS"
Write-Log "PSADT Installation Documentation Completed (Targeted)" "SUCCESS"
Write-Log "All documentation files saved to: $outputPath" "SUCCESS"
Write-Log "Log file saved to: $logPath" "SUCCESS"
Write-Log "========================================" "SUCCESS"
Write-Host "======================================"
Write-Host "Output folder: $outputPath" -ForegroundColor Cyan
Write-Host "`nGenerated Files:" -ForegroundColor Yellow
Write-Host "• JSON format for automation and uninstall/requirement script generation" -ForegroundColor White
Write-Host "• Complete execution log with detailed change analysis" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Copy the Documentation folder to your host system" -ForegroundColor Green
Write-Host "2. Use InstallationChanges JSON for automated uninstall and requirement script generation" -ForegroundColor Green
Write-Host "3. Review detailed log for installation analysis" -ForegroundColor Green

Write-Host "`nOpening documentation folder..." -ForegroundColor Cyan
Start-Process -FilePath "explorer.exe" -ArgumentList $outputPath

Write-Host "`n======================================"
Write-Host "AUTO-CLOSING SANDBOX IN 30 SECONDS" -ForegroundColor Red
Write-Host "======================================"
Write-Host "The Windows Sandbox will automatically close to clean up resources." -ForegroundColor Yellow
Write-Host "All documentation files have been saved to the mapped folder." -ForegroundColor Green
Write-Host "`nPress Ctrl+C to cancel auto-close..." -ForegroundColor Gray

# 30-second countdown with progress bar
for ($i = 30; $i -gt 0; $i--) {
    Write-Progress -Activity "Auto-closing Windows Sandbox" -Status "Sandbox will close automatically to free resources..." -SecondsRemaining $i -PercentComplete ((30 - $i) / 30 * 100)
    Start-Sleep -Seconds 1
}
Write-Progress -Activity "Auto-closing Windows Sandbox" -Completed

Write-Host "`nClosing Windows Sandbox..." -ForegroundColor Red
Write-Log "Sandbox auto-close initiated after 30-second countdown" "INFO"
Write-Log "Documentation process completed successfully" "SUCCESS"

# Gracefully shut down the sandbox so the host client can disconnect cleanly
Stop-Computer
'@

        # Replace placeholders with actual values
        $documentationScript = $documentationScript -replace '__PROJECTNAME__', $ProjectName
        $documentationScript = $documentationScript -replace '__TIMESTAMP__', $timestamp

        # Save the documentation script inside the PSADT project's SupportFiles folder
        $docScriptPath = Join-Path $supportFilesFolder "TargetedDocumentationScript.ps1"
        $documentationScript | Set-Content -Path $docScriptPath -Encoding UTF8
        Write-Host "✓ Targeted documentation script created: $docScriptPath" -ForegroundColor Green

        # Create the sandbox configuration file content
        $sandboxConfigContent = @"
<Configuration>
    <VGpu>Disable</VGpu>
    <Networking>Enable</Networking>
    <MappedFolders>
        <MappedFolder>
            <HostFolder>$ProjectPath</HostFolder>
            <SandboxFolder>C:\PSADT</SandboxFolder>
            <ReadOnly>false</ReadOnly>
        </MappedFolder>
    </MappedFolders>
    <LogonCommand>
        <Command>powershell.exe -NoExit -ExecutionPolicy Bypass -File C:\PSADT\SupportFiles\TargetedDocumentationScript.ps1</Command>
    </LogonCommand>
</Configuration>
"@

        # Write the sandbox configuration to the file
        $sandboxConfigContent | Set-Content -Path $sandboxConfigFile -Encoding UTF8
        Write-Host "✓ Targeted documentation sandbox configuration created!" -ForegroundColor Green

        # Check if Windows Sandbox is available
        try {
            $sandboxFeature = Get-WindowsOptionalFeature -Online -FeatureName "Containers-DisposableClientVM" -ErrorAction SilentlyContinue
            if ($sandboxFeature -and $sandboxFeature.State -ne "Enabled") {
                Write-Warning "Windows Sandbox feature is not enabled. Please enable it in Windows Features."
                Write-Host "To enable: Control Panel > Programs > Turn Windows features on or off > Windows Sandbox" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Unable to check Windows Sandbox feature status. Windows Sandbox may not be available on this system."
        }

        Write-Host "`nTargeted Documentation Setup Complete!" -ForegroundColor Cyan
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "Files created:" -ForegroundColor White
        Write-Host "• Sandbox Config: $sandboxConfigFile" -ForegroundColor Green
        Write-Host "• Documentation Script: $docScriptPath" -ForegroundColor Green
        
        # Launch Windows Sandbox automatically
        Write-Host "`nLaunching Windows Sandbox for targeted documentation..." -ForegroundColor Yellow
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "The sandbox will:" -ForegroundColor White
        Write-Host "• Scan targeted directories (Program Files, ProgramData, AppData)" -ForegroundColor Green
        Write-Host "• Monitor key registry locations (Uninstall, App Paths)" -ForegroundColor Green
        Write-Host "• Capture services and programs baseline" -ForegroundColor Green
        Write-Host "• Run your PSADT installation" -ForegroundColor Green
        Write-Host "• Compare before/after states to identify changes" -ForegroundColor Green
        Write-Host "• Generate detailed change reports" -ForegroundColor Green
        Write-Host "• Much faster than full system scans (2-5 minutes vs 10-30!)" -ForegroundColor Green
        Write-Host "===============================================" -ForegroundColor Cyan

        try {
            Start-Process -FilePath "WindowsSandbox.exe" -ArgumentList $sandboxConfigFile -ErrorAction Stop
            Write-Host "✓ Windows Sandbox launched successfully!" -ForegroundColor Green
            Write-Host "`nThe targeted documentation will be available in the sandbox at:" -ForegroundColor Cyan
            Write-Host "C:\PSADT\Documentation" -ForegroundColor White
        } catch {
            Write-Warning "Failed to start Windows Sandbox automatically: $($_.Exception.Message)"
            Write-Host "You can manually launch the sandbox by double-clicking: $sandboxConfigFile" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        Write-Warning "Failed to create targeted documentation: $($_.Exception.Message)"
        return $false
    }
}

function Wait-ForDocumentationAndProcess {
    param(
        [string]$ProjectPath,
        [string]$InstallerType
    )
    
    try {
        Write-Host "Monitoring for documentation completion..." -ForegroundColor Yellow
        
        # Look for JSON files in Documentation folder
        $docPath = Join-Path $ProjectPath "Documentation"
        $jsonFound = $false
        $jsonFile = $null
        $maxWaitMinutes = 30
        $checkIntervalSeconds = 10
        $totalChecks = ($maxWaitMinutes * 60) / $checkIntervalSeconds
        
        Write-Host "Checking for InstallationChanges JSON file every $checkIntervalSeconds seconds..." -ForegroundColor Cyan
        Write-Host "Maximum wait time: $maxWaitMinutes minutes" -ForegroundColor Gray
        
        for ($i = 1; $i -le $totalChecks; $i++) {
            if (Test-Path $docPath) {
                $jsonFiles = Get-ChildItem -Path $docPath -Filter "InstallationChanges*.json" -File -ErrorAction SilentlyContinue
                
                if ($jsonFiles.Count -gt 0) {
                    $jsonFile = $jsonFiles[0].FullName
                    $jsonFound = $true
                    Write-Host "✓ JSON documentation file found: $($jsonFiles[0].Name)" -ForegroundColor Green
                    break
                }
            }
            
            $minutesWaited = ($i * $checkIntervalSeconds) / 60
            Write-Host "Waiting... ($([math]::Round($minutesWaited, 1)) minutes elapsed)" -ForegroundColor Gray
            Start-Sleep -Seconds $checkIntervalSeconds
        }
        
        if (-not $jsonFound) {
            Write-Warning "Documentation JSON file not found after $maxWaitMinutes minutes. Please check the Windows Sandbox manually."
            return $false
        }
        
        Write-Host "`nProcessing documentation results..." -ForegroundColor Yellow
        
        # Generate requirement script
        Write-Host "Generating Intune requirement script..." -ForegroundColor Cyan
        $reqSuccess = New-IntuneRequirementScript -ProjectPath $ProjectPath -JsonFilePath $jsonFile
        
        if ($reqSuccess) {
            Write-Host "✓ Intune requirement script generated" -ForegroundColor Green
        } else {
            Write-Warning "Failed to generate requirement script"
        }
        
        # Generate uninstall logic (only for EXE installers, MSI uses Zero-Config)
        if ($InstallerType -eq 'exe') {
            Write-Host "Generating uninstall logic for EXE installer..." -ForegroundColor Cyan
            $uninstallSuccess = Update-PSADTUninstallLogic -ProjectPath $ProjectPath -JsonFilePath $jsonFile
            
            if ($uninstallSuccess) {
                Write-Host "✓ Uninstall logic generated for EXE installer" -ForegroundColor Green
            } else {
                Write-Warning "Failed to generate uninstall logic"
            }
        } else {
            Write-Host "✓ MSI installer detected - using Zero-Config uninstall (no manual logic needed)" -ForegroundColor Green
        }

        # Populate AppProcessesToClose for all installer types
        Write-Host "Detecting processes to close..." -ForegroundColor Cyan
        $procSuccess = Update-PSADTProcessesToClose -ProjectPath $ProjectPath -JsonFilePath $jsonFile
        if (-not $procSuccess) {
            Write-Warning "Could not auto-detect processes to close - AppProcessesToClose left empty"
        }
        
        return $true
    }
    catch {
        Write-Warning "Error processing documentation: $($_.Exception.Message)"
        return $false
    }
}

function New-IntuneRequirementScript {
    param(
        [string]$ProjectPath,
        [string]$JsonFilePath
    )
    
    try {
        # Parse JSON using exact logic from Create-IntuneRequirement.ps1
        Write-Host "Parsing JSON data..." -ForegroundColor White
        $jsonContent = Get-Content -Path $JsonFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
        
        # Extract application info using exact logic
        $appEntries = @()
        $productCodes = @()
        
        # Handle different JSON structures
        if ($jsonContent.Values -and $jsonContent.Values.PSObject.Properties) {
            # Direct Values structure
            foreach ($prop in $jsonContent.Values.PSObject.Properties) {
                $entry = $prop.Value
                if ($entry.DisplayName) {
                    $appEntries += $entry
                    if ($entry.UninstallString -and $entry.UninstallString -match '\{[A-F0-9-]{36}\}') {
                        $productCodes += $matches[0]
                    }
                }
            }
        } elseif ($jsonContent.NewRegistryKeys) {
            # NewRegistryKeys structure - search for Values within registry keys
            foreach ($regItem in $jsonContent.NewRegistryKeys) {
                if ($regItem.Values -and $regItem.Values.DisplayName) {
                    $entry = $regItem.Values
                    $appEntries += $entry
                    if ($entry.UninstallString -and $entry.UninstallString -match '\{[A-F0-9-]{36}\}') {
                        $productCodes += $matches[0]
                    }
                    if ($regItem.Path -and $regItem.Path -match '\{[A-F0-9-]{36}\}') {
                        $productCodes += $matches[0]
                    }
                }
            }
        } elseif ($jsonContent.NewPrograms -and $jsonContent.NewPrograms.Count -gt 0) {
            # NewPrograms array structure
            foreach ($prog in $jsonContent.NewPrograms) {
                if ($prog.DisplayName) {
                    $appEntries += $prog
                    if ($prog.UninstallString -and $prog.UninstallString -match '\{[A-F0-9-]{36}\}') {
                        $productCodes += $matches[0]
                    }
                }
            }
        }

        # If no entries found from JSON, fall back to the YAML manifest in the Files folder
        $appName = $null; $appVersion = $null; $publisher = $null
        if ($appEntries.Count -gt 0) {
            $mainApp    = $appEntries[0]
            $appName    = $mainApp.DisplayName
            $appVersion = $mainApp.DisplayVersion
            $publisher  = $mainApp.Publisher
            Write-Host "Extracted info for: $appName" -ForegroundColor Green
            if ($appVersion) { Write-Host "  Version: $appVersion" -ForegroundColor White }
            if ($publisher)  { Write-Host "  Publisher: $publisher" -ForegroundColor White }
            Write-Host "  Registry Entries: $($appEntries.Count)" -ForegroundColor White
            if ($productCodes.Count -gt 0) { Write-Host "  Product Codes: $($productCodes.Count)" -ForegroundColor White }
        } else {
            # Fallback: parse the winget YAML in Files\
            $yamlFile = Get-ChildItem (Join-Path $ProjectPath 'Files') -Filter '*.yaml' -ErrorAction SilentlyContinue |
                        Select-Object -First 1
            if ($yamlFile) {
                $yamlText = Get-Content $yamlFile.FullName -Raw
                if ($yamlText -match 'PackageName:\s*(.+)')    { $appName    = $Matches[1].Trim() }
                if ($yamlText -match 'PackageVersion:\s*(.+)') { $appVersion = $Matches[1].Trim() }
                if ($yamlText -match 'Publisher:\s*(.+)')      { $publisher  = $Matches[1].Trim() }
                if ($yamlText -match "ProductCode:\s*'?(\{[A-F0-9-]{36}\})'?") { $productCodes += $Matches[1] }
                if ($appName) {
                    Write-Host "JSON had no program entries — using YAML manifest as fallback" -ForegroundColor Yellow
                    Write-Host "  App: $appName  Version: $appVersion" -ForegroundColor White
                    if ($productCodes.Count -gt 0) { Write-Host "  Product Code: $($productCodes[0])" -ForegroundColor White }
                }
            }
            if (-not $appName) {
                Write-Warning "No application entries found in JSON file and no YAML fallback available"
                return $false
            }
        }
        
        # Generate requirement script using exact logic from original
        Write-Host "Generating requirement script..." -ForegroundColor White
        
        $requirementScript = @"
<#
.SYNOPSIS
    Intune Win32 App Requirement Script for $appName
.DESCRIPTION
    Checks if $appName is installed on the device.
    Generated automatically from InstallationChanges data on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
.NOTES
    This script should return exit code 0 if the requirement is met (app is installed)
    and exit code 1 if the requirement is not met (app is not installed or wrong version)
#>

try {
    `$appFound = `$false
    `$installedVersion = `$null
    `$requiredVersion = '$appVersion'
    
    # Registry paths to check for installed applications
    `$registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
"@

        # Add product code checks if available
        if ($productCodes.Count -gt 0) {
            $requirementScript += @"
    # Check specific product codes first
    `$productCodes = @(
"@
            foreach ($pc in $productCodes) {
                $requirementScript += "        '$pc'`n"
            }
            $requirementScript += @"
    )
    
    foreach (`$productCode in `$productCodes) {
        `$msiPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`$productCode"
        if (Test-Path `$msiPath) {
            `$msiApp = Get-ItemProperty -Path `$msiPath -ErrorAction SilentlyContinue
            if (`$msiApp -and `$msiApp.DisplayName -like '*$($appName.Split(' ')[0])*') {
                `$appFound = `$true
                `$installedVersion = `$msiApp.DisplayVersion
                Write-Host "Found via MSI Product Code: `$(`$msiApp.DisplayName) v`$installedVersion"
                break
            }
        }
    }
    
"@
        }

        # Add registry search for application name
        $requirementScript += @"
    # Search by application name if not found via product code
    if (-not `$appFound) {
        foreach (`$regPath in `$registryPaths) {
            try {
                `$apps = Get-ItemProperty -Path `$regPath -ErrorAction SilentlyContinue | Where-Object {
                    `$_.DisplayName -like '*$($appName.Split(' ')[0])*' -or
                    `$_.DisplayName -eq '$appName'
                }
                
                foreach (`$app in `$apps) {
                    if (`$app.DisplayName) {
                        `$appFound = `$true
                        `$installedVersion = `$app.DisplayVersion
                        Write-Host "Found application: `$(`$app.DisplayName)"
                        if (`$installedVersion) {
                            Write-Host "Installed Version: `$installedVersion"
                        }
"@

        # Add version comparison logic if we have a version to compare
        if ($appVersion) {
            $requirementScript += @"
                        
                        # Compare versions if available
                        if (`$installedVersion -and `$requiredVersion) {
                            try {
                                `$installedVer = [System.Version]`$installedVersion
                                `$requiredVer = [System.Version]`$requiredVersion
                                
                                if (`$installedVer -ge `$requiredVer) {
                                    Write-Host "Version requirement met: `$installedVersion >= `$requiredVersion"
                                    exit 0
                                } else {
                                    Write-Host "Version requirement NOT met: `$installedVersion < `$requiredVersion"
                                    exit 1
                                }
                            } catch {
                                # Version parsing failed, assume requirement is met if app is found
                                Write-Host "Version comparison failed, but application is installed"
                                exit 0
                            }
                        } else {
                            # No version info, just check if installed
                            Write-Host "Application found (no version comparison)"
                            exit 0
                        }
"@
        } else {
            $requirementScript += @"
                        
                        # No specific version required, app is installed
                        Write-Host "Application found"
                        exit 0
"@
        }

        $requirementScript += @"
                        break
                    }
                }
            } catch {
                # Continue checking other registry paths
                continue
            }
        }
    }
    
    # Final check - if app was found but no version comparison was done
    if (`$appFound) {
        Write-Host "Application is installed"
        exit 0
    } else {
        Write-Host "Application not found - requirement NOT met"
        exit 1
    }
    
} catch {
    Write-Host "Error during requirement check: `$(`$_.Exception.Message)"
    exit 1
}
"@

        # Save requirement script using exact logic from original
        $supportFilesPath = Join-Path $ProjectPath "SupportFiles"
        if (-not (Test-Path $supportFilesPath)) {
            New-Item -Path $supportFilesPath -ItemType Directory -Force | Out-Null
            Write-Host "Created SupportFiles directory" -ForegroundColor Yellow
        }
        
        $defaultPath = Join-Path $supportFilesPath "RequirementScript.ps1"
        
        $requirementScript | Set-Content -Path $defaultPath -Encoding UTF8
        Write-Host "`n✓ SUCCESS: Requirement script saved to:" -ForegroundColor Green
        Write-Host "  $defaultPath" -ForegroundColor White
        
        Write-Host "`nUsage Instructions:" -ForegroundColor Cyan
        Write-Host "1. Copy the generated PowerShell script content" -ForegroundColor White
        Write-Host "2. In Intune, go to your Win32 app > Requirements" -ForegroundColor White
        Write-Host "3. Add requirement rule: 'Script'" -ForegroundColor White
        Write-Host "4. Paste the script content" -ForegroundColor White
        Write-Host "5. Set 'Run script as 32-bit process': No" -ForegroundColor White
        Write-Host "6. Set 'Enforce script signature check': No" -ForegroundColor White
        
        return $true
    }
    catch {
        Write-Warning "Failed to create requirement script: $($_.Exception.Message)"
        return $false
    }
}

function Update-PSADTUninstallLogic {
    param(
        [string]$ProjectPath,
        [string]$JsonFilePath
    )
    
    try {
        $scriptPath = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
        
        if (-not (Test-Path $scriptPath)) {
            Write-Warning "PSADT script not found: $scriptPath"
            return $false
        }
        
        # Parse JSON data
        $jsonContent = Get-Content -Path $JsonFilePath -Raw -Encoding UTF8
        $data = $jsonContent | ConvertFrom-Json
        
        # Extract info using exact logic from Update-PSADTForUninstall-Working.ps1
        $appName = "Unknown App"
        $productCodes = @()
        $uninstallStrings = @()
        $installPaths = @()
        
        # Get app info from registry
        if ($data.NewRegistryKeys) {
            foreach ($regKey in $data.NewRegistryKeys) {
                if ($regKey.Path -like "*Uninstall*") {
                    # Handle the Values object structure
                    if ($regKey.Values) {
                        if ($regKey.Values.DisplayName) {
                            $appName = $regKey.Values.DisplayName
                        }
                        
                        # Smart uninstall string selection
                        $selectedUninstaller = $null
                        
                        # For EXE uninstallers: Prefer QuietUninstallString over UninstallString
                        if ($regKey.Values.QuietUninstallString -and $regKey.Values.QuietUninstallString -like "*.exe*") {
                            $selectedUninstaller = $regKey.Values.QuietUninstallString
                        }
                        elseif ($regKey.Values.UninstallString -and $regKey.Values.UninstallString -like "*.exe*") {
                            $selectedUninstaller = $regKey.Values.UninstallString
                        }
                        # For MSI uninstallers: Use UninstallString (contains msiexec)
                        elseif ($regKey.Values.UninstallString -and $regKey.Values.UninstallString -like "*msiexec*") {
                            $selectedUninstaller = $regKey.Values.UninstallString
                        }
                        
                        if ($selectedUninstaller) {
                            $uninstallStrings += $selectedUninstaller
                        }
                    }
                    
                    # Extract product codes
                    if ($regKey.Path -match '\{[A-F0-9-]{36}\}') {
                        $pc = [regex]::Match($regKey.Path, '\{[A-F0-9-]{36}\}').Value
                        if ($pc -notin $productCodes) {
                            $productCodes += $pc
                        }
                    }
                }
            }
        }
        
        # Get install paths from files
        if ($data.NewFiles) {
            foreach ($file in $data.NewFiles) {
                if ($file.Path -like "*Program Files*") {
                    $dir = Split-Path $file.Path -Parent
                    # Only include specific app directories, not system directories
                    if ($dir -notin $installPaths -and $dir -notlike "*Program Files" -and $dir -notlike "*ProgramData" -and $dir -ne "C:\Program Files" -and $dir -ne "C:\ProgramData") {
                        $installPaths += $dir
                    }
                }
            }
        }
        
        Write-Host "Extracted info for: $appName" -ForegroundColor Green
        Write-Host "  Product Codes: $($productCodes.Count)" -ForegroundColor White
        Write-Host "  Uninstall Strings: $($uninstallStrings.Count)" -ForegroundColor White
        Write-Host "  Install Paths: $($installPaths.Count)" -ForegroundColor White
        
        # Generate uninstall code using exact logic from working script
        $codeLines = @()
        $codeLines += "        # Auto-generated uninstall code for $appName"
        $codeLines += "        # Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $codeLines += ""
        $codeLines += "        Write-ADTLogEntry -Message `"Starting uninstall of $appName`""
        $codeLines += "        `$uninstallSuccess = `$false"
        $codeLines += ""
        
        # Add MSI uninstall if product codes found
        if ($productCodes.Count -gt 0) {
            $codeLines += "        # Try MSI uninstall"
            $codeLines += "        if (-not `$uninstallSuccess) {"
            $codeLines += "            try {"
            $codeLines += "                Write-ADTLogEntry -Message `"Attempting MSI uninstallation`""
            
            foreach ($pc in $productCodes) {
                $codeLines += "                Write-ADTLogEntry -Message `"Uninstalling MSI with Product Code: $pc`""
                $codeLines += "                `$result = Start-ADTMsiProcess -Action 'Uninstall' -ProductCode '$pc' -PassThru"
                $codeLines += "                `$exitCode = `$result.ExitCode"
                $codeLines += "                if (`$exitCode -eq 0 -or `$exitCode -eq 3010) {"
                $codeLines += "                    Write-ADTLogEntry -Message `"MSI uninstallation completed successfully (Exit Code: `$exitCode)`""
                $codeLines += "                    `$uninstallSuccess = `$true"
                $codeLines += "                    break"
                $codeLines += "                } else {"
                $codeLines += "                    Write-ADTLogEntry -Message `"MSI uninstallation failed with exit code: `$exitCode`" -Severity 2"
                $codeLines += "                }"
            }
            
            $codeLines += "            } catch {"
            $codeLines += "                Write-ADTLogEntry -Message `"MSI uninstall failed: `$(`$_.Exception.Message)`" -Severity 2"
            $codeLines += "            }"
            $codeLines += "        }"
            $codeLines += ""
        }
        
        # Process registry uninstall strings and determine installer type
        if ($uninstallStrings.Count -gt 0) {
            foreach ($us in $uninstallStrings) {
                # Determine installer type and generate appropriate code
                if ($us -like '*msiexec*') {
                    # MSI-based uninstaller - extract product code
                    if ($us -match '\{[A-F0-9-]{36}\}') {
                        $productCode = $matches[0]
                        $codeLines += "        # Try MSI uninstall with product code"
                        $codeLines += "        if (-not `$uninstallSuccess) {"
                        $codeLines += "            try {"
                        $codeLines += "                Write-ADTLogEntry -Message `"Attempting MSI uninstallation with product code: $productCode`""
                        $codeLines += "                `$result = Start-ADTMsiProcess -Action 'Uninstall' -ProductCode '$productCode' -PassThru"
                        $codeLines += "                `$exitCode = `$result.ExitCode"
                        $codeLines += "                if (`$exitCode -eq 0 -or `$exitCode -eq 3010) {"
                        $codeLines += "                    Write-ADTLogEntry -Message `"MSI uninstallation completed successfully (Exit Code: `$exitCode)`""
                        $codeLines += "                    `$uninstallSuccess = `$true"
                        $codeLines += "                } else {"
                        $codeLines += "                    Write-ADTLogEntry -Message `"MSI uninstallation failed with exit code: `$exitCode`" -Severity 2"
                        $codeLines += "                }"
                        $codeLines += "            } catch {"
                        $codeLines += "                Write-ADTLogEntry -Message `"MSI uninstall failed: `$(`$_.Exception.Message)`" -Severity 2"
                        $codeLines += "            }"
                        $codeLines += "        }"
                        $codeLines += ""
                    }
                }
                elseif ($us -like '*.exe*') {
                    # EXE-based uninstaller - parse path and parameters
                    $exePath = ""
                    $exeParams = ""
                    
                    if ($us -match '"([^"]*\.exe)"(.*)') {
                        $exePath = $matches[1]
                        $exeParams = $matches[2].Trim()
                    } elseif ($us -match '([^\s]*\.exe)(.*)') {
                        $exePath = $matches[1]
                        $exeParams = $matches[2].Trim()
                    } else {
                        $exePath = $us
                    }
                    
                    $codeLines += "        # Try EXE uninstall"
                    $codeLines += "        if (-not `$uninstallSuccess) {"
                    $codeLines += "            try {"
                    $codeLines += "                Write-ADTLogEntry -Message `"Attempting EXE uninstallation: $exePath`""
                    $codeLines += "                `$result = Start-ADTProcess -FilePath '$exePath' -ArgumentList '$exeParams' -PassThru"
                    $codeLines += "                `$exitCode = `$result.ExitCode"
                    $codeLines += "                if (`$exitCode -eq 0 -or `$exitCode -eq 3010) {"
                    $codeLines += "                    Write-ADTLogEntry -Message `"EXE uninstallation completed successfully (Exit Code: `$exitCode)`""
                    $codeLines += "                    `$uninstallSuccess = `$true"
                    $codeLines += "                } else {"
                    $codeLines += "                    Write-ADTLogEntry -Message `"EXE uninstallation failed with exit code: `$exitCode`" -Severity 2"
                    $codeLines += "                }"
                    $codeLines += "            } catch {"
                    $codeLines += "                Write-ADTLogEntry -Message `"EXE uninstall failed: `$(`$_.Exception.Message)`" -Severity 2"
                    $codeLines += "            }"
                    $codeLines += "        }"
                    $codeLines += ""
                }
            }
        }
        
        # Add cleanup
        if ($installPaths.Count -gt 0) {
            $codeLines += "        # Cleanup installation files"
            $codeLines += "        Write-ADTLogEntry -Message `"Performing cleanup of installation directories`""
            
            foreach ($path in $installPaths) {
                $codeLines += "        if (Test-Path '$path') {"
                $codeLines += "            try {"
                $codeLines += "                Write-ADTLogEntry -Message `"Removing directory: $path`""
                $codeLines += "                Remove-ADTFolder -Path '$path'"
                $codeLines += "                Write-ADTLogEntry -Message `"Successfully removed directory: $path`""
                $codeLines += "            } catch {"
                $codeLines += "                Write-ADTLogEntry -Message `"Failed to remove directory $path`: `$(`$_.Exception.Message)`" -Severity 2"
                $codeLines += "            }"
                $codeLines += "        } else {"
                $codeLines += "            Write-ADTLogEntry -Message `"Directory not found (already removed): $path`""
                $codeLines += "        }"
            }
            $codeLines += ""
        }
        
        $codeLines += "        # Verify uninstallation"
        $codeLines += "        if (`$uninstallSuccess) {"
        $codeLines += "            Write-ADTLogEntry -Message `"Uninstallation completed successfully for $appName`""
        $codeLines += "        } else {"
        $codeLines += "            Write-ADTLogEntry -Message `"Warning: Uninstallation may not have completed successfully - manual verification recommended`" -Severity 2"
        $codeLines += "        }"
        $codeLines += ""
        
        $uninstallCode = $codeLines -join "`r`n"
        
        # Update PSADT file using exact logic from working script
        $content = Get-Content -Path $scriptPath -Raw -Encoding UTF8
        
        # Find and replace the correct uninstall section
        $uninstallMarker = '## <Perform Uninstallation tasks here>'
        
        if ($content -match [regex]::Escape($uninstallMarker)) {
            # Find the exact position to insert the code
            $lines = $content -split "`r?`n"
            $markerIndex = -1
            
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i].Trim() -eq $uninstallMarker) {
                    $markerIndex = $i
                    break
                }
            }
            
            if ($markerIndex -ne -1) {
                # Insert the uninstall code after the marker
                $beforeMarker = $lines[0..$markerIndex]
                $afterMarker = $lines[($markerIndex + 1)..($lines.Count - 1)]
                
                # Clean up the generated code (remove any header comments)
                $cleanCode = $uninstallCode -split "`r?`n" | Where-Object { 
                    $_ -notmatch '^\s*##\*' -and 
                    $_ -notmatch '^\s*\[String\]\$installPhase' 
                }
                
                # Combine all parts
                $newLines = @()
                $newLines += $beforeMarker
                $newLines += ""
                $newLines += $cleanCode
                $newLines += $afterMarker
                
                $content = $newLines -join "`r`n"
            } else {
                throw "Could not find the uninstall marker in PSADT file."
            }
        } else {
            throw "Could not find uninstall section in PSADT file. Please ensure the file contains '## <Perform Uninstallation tasks here>' marker."
        }
        
        $content | Set-Content -Path $scriptPath -Encoding UTF8
        
        Write-Host "✓ SUCCESS: Updated PSADT for $appName" -ForegroundColor Green
        Write-Host "Methods: $(if($productCodes.Count -gt 0){'MSI '})$(if($uninstallStrings.Count -gt 0){'Registry '})Cleanup" -ForegroundColor White
        return $true
        
    } catch {
        Write-Warning "Failed to update uninstall logic: $($_.Exception.Message)"
        return $false
    }
}

function Update-PSADTProcessesToClose {
    param(
        [string]$ProjectPath,
        [string]$JsonFilePath
    )

    try {
        $scriptPath = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
        if (-not (Test-Path $scriptPath)) {
            Write-Warning "PSADT script not found: $scriptPath"
            return $false
        }

        $data = Get-Content -Path $JsonFilePath -Raw -Encoding UTF8 | ConvertFrom-Json

        $candidates  = [System.Collections.Generic.List[string]]::new()
        $excludePattern = 'uninstall|uninst|setup|install|update|patch|redist'

        # Get InstallLocation from Uninstall registry key (used as scope filter for fallback)
        $installLocation = $null
        foreach ($regKey in $data.NewRegistryKeys) {
            if ($regKey.Path -like '*Uninstall*' -and $regKey.Values -and $regKey.Values.InstallLocation) {
                $installLocation = $regKey.Values.InstallLocation.TrimEnd('\')
                break
            }
        }

        # Source 1 (preferred): App Paths registry keys — these are user-launchable executables
        foreach ($regKey in $data.NewRegistryKeys) {
            if ($regKey.Path -like '*App Paths*' -and $regKey.KeyName -like '*.exe') {
                $procName = [System.IO.Path]::GetFileNameWithoutExtension($regKey.KeyName)
                if ($procName -and $procName -notmatch $excludePattern -and $procName -notin $candidates) {
                    $candidates.Add($procName)
                }
            }
        }

        # Source 2 (fallback): DisplayIcon from Uninstall registry key
        if ($candidates.Count -eq 0) {
            foreach ($regKey in $data.NewRegistryKeys) {
                if ($regKey.Path -like '*Uninstall*' -and $regKey.Values -and $regKey.Values.DisplayIcon) {
                    $iconPath = $regKey.Values.DisplayIcon -replace ',\d+$', ''  # strip icon index
                    $procName = [System.IO.Path]::GetFileNameWithoutExtension($iconPath)
                    if ($procName -and $procName -notmatch $excludePattern -and $procName -notin $candidates) {
                        $candidates.Add($procName)
                    }
                    break
                }
            }
        }

        # Source 3 (fallback): EXE files under InstallLocation
        if ($candidates.Count -eq 0 -and $installLocation) {
            foreach ($file in $data.NewFiles) {
                if ($file.Type -eq 'File' -and
                    $file.Path -like '*.exe' -and
                    $file.Path.StartsWith($installLocation, [System.StringComparison]::OrdinalIgnoreCase)) {
                    # Only files directly in InstallLocation, not subfolders
                    $fileDir = [System.IO.Path]::GetDirectoryName($file.Path)
                    if ($fileDir -ieq $installLocation) {
                        $procName = [System.IO.Path]::GetFileNameWithoutExtension($file.Path)
                        if ($procName -and $procName -notmatch $excludePattern -and $procName -notin $candidates) {
                            $candidates.Add($procName)
                        }
                    }
                }
            }
        }

        # Sort and build the replacement string
        $sorted = $candidates | Sort-Object
        $arrayContent = if ($sorted.Count -gt 0) {
            ($sorted | ForEach-Object { "'$_'" }) -join ', '
        } else { '' }
        $replacement = "AppProcessesToClose = @($arrayContent)"

        # Regex-replace the existing AppProcessesToClose line (idempotent)
        $content = Get-Content -Path $scriptPath -Raw -Encoding UTF8
        $content = $content -replace 'AppProcessesToClose\s*=\s*@\([^)]*\)', $replacement
        Set-Content -Path $scriptPath -Value $content -Encoding UTF8

        if ($sorted.Count -gt 0) {
            Write-Host "✓ AppProcessesToClose set: @($arrayContent)" -ForegroundColor Green
        } else {
            Write-Host "AppProcessesToClose: no user-launchable processes detected, left as @()" -ForegroundColor DarkYellow
        }

        return $true
    }
    catch {
        Write-Warning "Failed to update AppProcessesToClose: $($_.Exception.Message)"
        return $false
    }
}

#region ── Org Template functions ────────────────────────────────────────────

# Current version of the template schema. Bump this when new fields are added.
$script:TemplateSchemaVersion = '2.0'

function New-OrgTemplate {
    param([PSCustomObject]$ExistingTemplate = $null)

    $templateFolder = Join-Path $env:APPDATA 'IntuneToolkit'
    if (-not (Test-Path $templateFolder)) {
        New-Item -ItemType Directory -Path $templateFolder -Force | Out-Null
    }

    # Detect installed PSADT version for embedding in template
    $psadtVer = (Get-Module -Name PSAppDeployToolkit -ListAvailable |
        Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()
    if (-not $psadtVer) { $psadtVer = 'unknown' }

    Write-Host ''
    Write-Host '=== Organisation Template Wizard ===' -ForegroundColor Cyan
    Write-Host 'Only Fluent dialog style is supported.' -ForegroundColor DarkGray
    Write-Host 'Press Enter on any prompt to keep the default value.' -ForegroundColor DarkGray
    Write-Host ''

    function Read-TV { param([string]$P, [string]$D='') $s=if($D){"[$D] "}else{''}; $v=Read-Host "$P $s"; if([string]::IsNullOrWhiteSpace($v)){$D}else{$v.Trim()} }
    function Read-TB { param([string]$P, [bool]$D) $d=if($D){'y'}else{'n'}; $v=Read-Host "$P (y/n) [$d]"; if([string]::IsNullOrWhiteSpace($v)){$D}else{$v.Trim() -in @('y','Y','yes','Yes','1','true','True')} }
    function Read-TI { param([string]$P, [int]$D) $v=Read-Host "$P [$D]"; if([string]::IsNullOrWhiteSpace($v)){$D}else{$n=0;if([int]::TryParse($v.Trim(),[ref]$n)){$n}else{$D}} }

    Write-Host '--- A: Identity ---' -ForegroundColor Yellow
    $templateName = Read-TV 'Template name'                                    ($ExistingTemplate?.TemplateName    ?? 'Default')
    $companyName  = Read-TV 'Company name (shown in dialog subtitles)'         ($ExistingTemplate?.CompanyName     ?? 'Your Organisation IT')
    $scriptAuthor = Read-TV 'App script author'                                ($ExistingTemplate?.AppScriptAuthor ?? 'IT Packaging Team')

    Write-Host ''; Write-Host '--- B: Branding (Fluent only) ---' -ForegroundColor Yellow
    $accentColor = Read-TV 'Fluent accent hex e.g. 0xFF0078D7 (blank=default)' ($ExistingTemplate?.FluentAccentColor ?? '')
    $logPath     = Read-TV 'Log path'                                           ($ExistingTemplate?.LogPath           ?? '$envWinDir\Logs\Software')

    Write-Host ''; Write-Host '--- C: Progress Messages ---' -ForegroundColor Yellow
    $pMsgI = Read-TV 'Progress message - Install'   ($ExistingTemplate?.ProgressMessage?.Install   ?? 'Installation in progress. Please wait...')
    $pMsgR = Read-TV 'Progress message - Repair'    ($ExistingTemplate?.ProgressMessage?.Repair    ?? 'Repair in progress. Please wait...')
    $pMsgU = Read-TV 'Progress message - Uninstall' ($ExistingTemplate?.ProgressMessage?.Uninstall ?? 'Uninstallation in progress. Please wait...')
    $pDtlI = Read-TV 'Progress detail - Install'    ($ExistingTemplate?.ProgressMessageDetail?.Install   ?? 'This window will close automatically when the installation is complete.')
    $pDtlR = Read-TV 'Progress detail - Repair'     ($ExistingTemplate?.ProgressMessageDetail?.Repair    ?? 'This window will close automatically when the repair is complete.')
    $pDtlU = Read-TV 'Progress detail - Uninstall'  ($ExistingTemplate?.ProgressMessageDetail?.Uninstall ?? 'This window will close automatically when the uninstallation is complete.')

    Write-Host ''; Write-Host '--- C: Balloon Notifications ---' -ForegroundColor Yellow
    $balI = Read-TV 'Balloon complete - Install'   ($ExistingTemplate?.BalloonComplete?.Install   ?? 'Installation complete.')
    $balR = Read-TV 'Balloon complete - Repair'    ($ExistingTemplate?.BalloonComplete?.Repair    ?? 'Repair complete.')
    $balU = Read-TV 'Balloon complete - Uninstall' ($ExistingTemplate?.BalloonComplete?.Uninstall ?? 'Uninstallation complete.')

    Write-Host ''; Write-Host '--- D: Install Welcome Dialog ---' -ForegroundColor Yellow
    $wEnabled    = Read-TB 'Show Welcome dialog on install'              ($ExistingTemplate?.WelcomeDialog?.Enabled              ?? $true)
    $wDefer      = Read-TB 'Allow deferral'                              ($ExistingTemplate?.WelcomeDialog?.AllowDefer           ?? $true)
    $wDeferTimes = Read-TI 'Max deferrals'                               ($ExistingTemplate?.WelcomeDialog?.DeferTimes           ?? 3)
    $wDisk       = Read-TB 'Check disk space'                            ($ExistingTemplate?.WelcomeDialog?.CheckDiskSpace       ?? $false)
    $wPersist    = Read-TB 'Persist prompt (user cannot ignore)'         ($ExistingTemplate?.WelcomeDialog?.PersistPrompt        ?? $true)
    $wBlock      = Read-TB 'Block app re-launch during install'          ($ExistingTemplate?.WelcomeDialog?.BlockExecution       ?? $true)
    $wCountdown  = Read-TI 'Auto-close countdown seconds (0=off)'        ($ExistingTemplate?.WelcomeDialog?.CloseProcessesCountdown ?? 180)
    $wCustomText = Read-TB 'Show custom text from strings.psd1'          ($ExistingTemplate?.WelcomeDialog?.CustomText           ?? $false)

    Write-Host ''; Write-Host '--- D: Uninstall Welcome Dialog ---' -ForegroundColor Yellow
    Write-Host '  (shown only when processes need closing before uninstall)' -ForegroundColor DarkGray
    $uwEnabled   = Read-TB 'Show Welcome dialog on uninstall'            ($ExistingTemplate?.UninstallWelcomeDialog?.Enabled              ?? $true)
    $uwCountdown = Read-TI 'Auto-close countdown seconds (0=off)'        ($ExistingTemplate?.UninstallWelcomeDialog?.CloseProcessesCountdown ?? 60)
    $uwPersist   = Read-TB 'Persist prompt'                              ($ExistingTemplate?.UninstallWelcomeDialog?.PersistPrompt         ?? $false)
    $uwBlock     = Read-TB 'Block app re-launch during uninstall'        ($ExistingTemplate?.UninstallWelcomeDialog?.BlockExecution        ?? $false)

    Write-Host ''; Write-Host '--- D: Progress Dialog ---' -ForegroundColor Yellow
    $prEnabled = Read-TB  'Show Progress dialog'                                         ($ExistingTemplate?.ProgressDialog?.Enabled             ?? $true)
    $prStatus  = Read-TV  'Override status message (blank = use strings.psd1)'           ($ExistingTemplate?.ProgressDialog?.StatusMessage       ?? '')
    $prDetail  = Read-TV  'Override detail text, Fluent only (blank = use strings.psd1)' ($ExistingTemplate?.ProgressDialog?.StatusMessageDetail ?? '')

    Write-Host ''; Write-Host '--- D: Completion Prompt (Post-Install) ---' -ForegroundColor Yellow
    $cpEnabled = Read-TB 'Show completion prompt after install'  ($ExistingTemplate?.CompletionPrompt?.Enabled         ?? $false)
    $cpMessage = Read-TV 'Completion message'                   ($ExistingTemplate?.CompletionPrompt?.Message         ?? 'The installation has completed successfully.')
    $cpButton  = Read-TV 'Button label'                         ($ExistingTemplate?.CompletionPrompt?.ButtonRightText ?? 'OK')

    $template = [PSCustomObject]@{
        TemplateSchemaVersion = $script:TemplateSchemaVersion
        TemplateName          = $templateName
        CompanyName           = $companyName
        AppScriptAuthor       = $scriptAuthor
        DialogStyle           = 'Fluent'
        FluentAccentColor     = $accentColor
        LogPath               = $logPath
        PsadtVersion          = $psadtVer
        ProgressMessage       = [PSCustomObject]@{ Install = $pMsgI; Repair = $pMsgR; Uninstall = $pMsgU }
        ProgressMessageDetail = [PSCustomObject]@{ Install = $pDtlI; Repair = $pDtlR; Uninstall = $pDtlU }
        BalloonComplete       = [PSCustomObject]@{ Install = $balI; Repair = $balR; Uninstall = $balU }
        WelcomeDialog         = [PSCustomObject]@{
            Enabled                 = $wEnabled
            AllowDefer              = $wDefer
            DeferTimes              = $wDeferTimes
            CheckDiskSpace          = $wDisk
            PersistPrompt           = $wPersist
            BlockExecution          = $wBlock
            CloseProcessesCountdown = $wCountdown
            CustomText              = $wCustomText
        }
        UninstallWelcomeDialog = [PSCustomObject]@{
            Enabled                 = $uwEnabled
            CloseProcessesCountdown = $uwCountdown
            PersistPrompt           = $uwPersist
            BlockExecution          = $uwBlock
        }
        ProgressDialog        = [PSCustomObject]@{
            Enabled             = $prEnabled
            StatusMessage       = $prStatus
            StatusMessageDetail = $prDetail
        }
        CompletionPrompt      = [PSCustomObject]@{
            Enabled         = $cpEnabled
            Message         = $cpMessage
            ButtonRightText = $cpButton
        }
    }

    $savePath = Join-Path $templateFolder "$templateName.json"
    $template | ConvertTo-Json -Depth 10 | Set-Content -Path $savePath -Encoding UTF8
    Write-Host ''
    Write-Host "✓ Template '$templateName' saved: $savePath" -ForegroundColor Green
    return $template
}

function Update-OrgTemplateIfNeeded {
    param(
        [PSCustomObject]$Template,
        [string]$FilePath
    )

    $schemaVer   = if ($Template.TemplateSchemaVersion) { $Template.TemplateSchemaVersion } else { '1.0' }
    $currentVer  = $script:TemplateSchemaVersion
    $needsUpdate = ([System.Version]$schemaVer) -lt ([System.Version]$currentVer)

    if (-not $needsUpdate) { return $Template }

    Write-Host ''
    Write-Host "  Template '$($Template.TemplateName)' uses schema v$schemaVer, current is v$currentVer." -ForegroundColor Yellow
    Write-Host '  New options are available. Would you like to update the template now?' -ForegroundColor Yellow
    $ans = Read-Host '  Update template? (y/n) [y]'
    if ($ans.Trim() -in @('','y','Y','yes','Yes','1','true')) {
        Write-Host '  Opening template wizard with existing values pre-filled...' -ForegroundColor Cyan
        $updated = New-OrgTemplate -ExistingTemplate $Template
        return $updated
    }

    # User skipped — patch missing fields silently with defaults so Apply-OrgTemplate won't fail
    if (-not $Template.PSObject.Properties['TemplateSchemaVersion']) { $Template | Add-Member -NotePropertyName TemplateSchemaVersion -NotePropertyValue $currentVer -Force }
    if (-not $Template.PSObject.Properties['PsadtVersion'])          { $Template | Add-Member -NotePropertyName PsadtVersion -NotePropertyValue 'unknown' -Force }
    if (-not $Template.PSObject.Properties['DialogStyle'])           { $Template | Add-Member -NotePropertyName DialogStyle -NotePropertyValue 'Fluent' -Force }
    if (-not $Template.PSObject.Properties['UninstallWelcomeDialog']) {
        $Template | Add-Member -NotePropertyName UninstallWelcomeDialog -NotePropertyValue (
            [PSCustomObject]@{ Enabled = $true; CloseProcessesCountdown = 60; PersistPrompt = $false; BlockExecution = $false }
        ) -Force
    }
    return $Template
}

function Get-OrgTemplate {
    $templateFolder = Join-Path $env:APPDATA 'IntuneToolkit'
    if (-not (Test-Path $templateFolder)) {
        New-Item -ItemType Directory -Path $templateFolder -Force | Out-Null
    }

    $templates = @(Get-ChildItem -Path $templateFolder -Filter '*.json' -ErrorAction SilentlyContinue)

    if ($templates.Count -eq 0) {
        Write-Host ''
        Write-Host 'No org template found. Creating one now...' -ForegroundColor Cyan
        return New-OrgTemplate
    }

    if ($templates.Count -eq 1) {
        $t = Get-Content -Path $templates[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $t = Update-OrgTemplateIfNeeded -Template $t -FilePath $templates[0].FullName
        Write-Host "✓ Loaded org template: $($t.TemplateName)" -ForegroundColor Green
        return $t
    }

    # Multiple templates — show picker
    Write-Host ''
    Write-Host '  Multiple org templates found:' -ForegroundColor Cyan
    Write-Host ("  {0,-4} {1}" -f '#', 'Template') -ForegroundColor Gray
    Write-Host ("  " + "-" * 40) -ForegroundColor DarkGray

    for ($i = 0; $i -lt $templates.Count; $i++) {
        try { $label = (Get-Content -Path $templates[$i].FullName -Raw -Encoding UTF8 | ConvertFrom-Json).TemplateName }
        catch { $label = $templates[$i].BaseName }
        if (-not $label) { $label = $templates[$i].BaseName }
        $color = if ($i % 2 -eq 0) { 'Cyan' } else { 'White' }
        Write-Host ("  {0,-4} {1}" -f ($i + 1), $label) -ForegroundColor $color
    }
    $newIdx = $templates.Count + 1
    $editIdx = $templates.Count + 2
    Write-Host ("  {0,-4} {1}" -f $newIdx,  '[Create new template]') -ForegroundColor DarkGray
    Write-Host ("  {0,-4} {1}" -f $editIdx, '[Edit existing template]') -ForegroundColor DarkGray
    Write-Host ''

    do {
        $rawInput = Read-Host "Select template (1-$editIdx)"
        $parsed   = 0
        $valid    = [int]::TryParse($rawInput.Trim(), [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $editIdx
        if (-not $valid) { Write-Host "Please enter a number between 1 and $editIdx." -ForegroundColor Red }
    } while (-not $valid)

    if ($parsed -eq $newIdx)  { return New-OrgTemplate }
    if ($parsed -eq $editIdx) {
        # Pick which one to edit then open wizard pre-filled
        do {
            $rawInput2 = Read-Host "Select template to edit (1-$($templates.Count))"
            $parsed2   = 0
            $v2 = [int]::TryParse($rawInput2.Trim(), [ref]$parsed2) -and $parsed2 -ge 1 -and $parsed2 -le $templates.Count
            if (-not $v2) { Write-Host "Please enter a number between 1 and $($templates.Count)." -ForegroundColor Red }
        } while (-not $v2)
        $existing = Get-Content -Path $templates[$parsed2 - 1].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        return New-OrgTemplate -ExistingTemplate $existing
    }

    $t = Get-Content -Path $templates[$parsed - 1].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $t = Update-OrgTemplateIfNeeded -Template $t -FilePath $templates[$parsed - 1].FullName
    Write-Host "✓ Loaded org template: $($t.TemplateName)" -ForegroundColor Green
    return $t
}

function Apply-OrgTemplate {
    param(
        [string]$ProjectPath,
        [PSCustomObject]$Template
    )

    try {
        # Helper: replace first regex match in string using index maths (avoids $ capture-group issues)
        function Set-TextBlock {
            param([string]$Text, [string]$Pattern, [string]$Replacement, [switch]$Multiline)
            $opts = if ($Multiline) { [System.Text.RegularExpressions.RegexOptions]::Singleline } `
                                    else { [System.Text.RegularExpressions.RegexOptions]::None }
            $m = [regex]::Match($Text, $Pattern, $opts)
            if (-not $m.Success) { return $Text }
            $Text.Substring(0, $m.Index) + $Replacement + $Text.Substring($m.Index + $m.Length)
        }

        #── config.psd1 ────────────────────────────────────────────────────
        $configPath = Join-Path $ProjectPath 'Config\config.psd1'
        if (Test-Path $configPath) {
            $cfg = Get-Content $configPath -Raw -Encoding UTF8

            $cfg = $cfg -replace "CompanyName = '[^']*'",        "CompanyName = '$($Template.CompanyName)'"
            $cfg = $cfg -replace "DialogStyle = '(Fluent|Classic)'", "DialogStyle = '$($Template.DialogStyle)'"

            if ($Template.FluentAccentColor -and $Template.FluentAccentColor.Trim() -ne '') {
                $cfg = $cfg -replace 'FluentAccentColor = [^\r\n]+', "FluentAccentColor = $($Template.FluentAccentColor)"
            } else {
                $cfg = $cfg -replace 'FluentAccentColor = [^\r\n]+', 'FluentAccentColor = $null'
            }

            if ($Template.LogPath -and $Template.LogPath.Trim() -ne '') {
                # Only target the Toolkit.LogPath (line after its specific comment)
                $cfg = $cfg -replace '(?m)(# Log path used for Toolkit logging\.\r?\n\s*LogPath = )[^\r\n]+', "`${1}'$($Template.LogPath)'"
            }

            Set-Content -Path $configPath -Value $cfg -Encoding UTF8
            Write-Host '  ✓ config.psd1 updated' -ForegroundColor Green
        }

        #── strings.psd1 ───────────────────────────────────────────────────
        $strPath = Join-Path $ProjectPath 'Strings\strings.psd1'
        if (Test-Path $strPath) {
            $str = Get-Content $strPath -Raw -Encoding UTF8

            # BalloonTip.Complete sub-block
            $newBalloon = "        # Text displayed in the balloon tip for successful completion of a deployment type.`r`n        Complete = @{`r`n            Install = '$($Template.BalloonComplete.Install)'`r`n            Repair = '$($Template.BalloonComplete.Repair)'`r`n            Uninstall = '$($Template.BalloonComplete.Uninstall)'`r`n        }"
            $str = Set-TextBlock -Text $str `
                -Pattern '(?s)        # Text displayed in the balloon tip for successful completion of a deployment type\.\r?\n        Complete = @\{[^}]*\}' `
                -Replacement $newBalloon -Multiline

            # ProgressPrompt — entire block up to RestartPrompt
            $nl = if ($str -match '\r\n') { "`r`n" } else { "`n" }
            $newProg = "    ProgressPrompt = @{${nl}        # Default message displayed in the progress bar.${nl}        Message = @{${nl}            Install = '$($Template.ProgressMessage.Install)'${nl}            Repair = '$($Template.ProgressMessage.Repair)'${nl}            Uninstall = '$($Template.ProgressMessage.Uninstall)'${nl}        }${nl}${nl}        # Default message detail displayed in the progress bar.${nl}        MessageDetail = @{${nl}            Install = '$($Template.ProgressMessageDetail.Install)'${nl}            Repair = '$($Template.ProgressMessageDetail.Repair)'${nl}            Uninstall = '$($Template.ProgressMessageDetail.Uninstall)'${nl}        }${nl}${nl}        # The subtitle underneath the Install Title, e.g. Company Name. Only for Fluent dialogs.${nl}        Subtitle = @{${nl}            Install = '{Toolkit\CompanyName} - App Installation'${nl}            Repair = '{Toolkit\CompanyName} - App Repair'${nl}            Uninstall = '{Toolkit\CompanyName} - App Uninstallation'${nl}        }${nl}    }"
            $str = Set-TextBlock -Text $str `
                -Pattern '(?s)    ProgressPrompt = @\{.*?(?=\r?\n    RestartPrompt)' `
                -Replacement $newProg -Multiline

            Set-Content -Path $strPath -Value $str -Encoding UTF8
            Write-Host '  ✓ strings.psd1 updated' -ForegroundColor Green
        }

        #── Invoke-AppDeployToolkit.ps1 ─────────────────────────────────────
        $scriptPath = Join-Path $ProjectPath 'Invoke-AppDeployToolkit.ps1'
        if (Test-Path $scriptPath) {
            $scr = Get-Content $scriptPath -Raw -Encoding UTF8

            # AppScriptAuthor
            $scr = $scr -replace "AppScriptAuthor = '[^']*'", "AppScriptAuthor = '$($Template.AppScriptAuthor)'"

            # ── Install Welcome dialog ──
            if ($Template.WelcomeDialog.Enabled) {
                $saiwLines       = [System.Collections.Generic.List[string]]::new()  # base hashtable entries
                $saiwCloseLines  = [System.Collections.Generic.List[string]]::new()  # only valid with CloseProcesses
                if ($Template.WelcomeDialog.AllowDefer) {
                    $saiwLines.Add("        AllowDefer = `$true")
                    $saiwLines.Add("        DeferTimes = $($Template.WelcomeDialog.DeferTimes)")
                }
                if ($Template.WelcomeDialog.CheckDiskSpace) { $saiwLines.Add("        CheckDiskSpace = `$true") }
                if ($Template.WelcomeDialog.PersistPrompt)  { $saiwLines.Add("        PersistPrompt = `$true") }
                if ($Template.WelcomeDialog.CustomText) {
                    # Only emit CustomText if strings.psd1 actually has a non-empty CustomMessage
                    $strPathForCheck = Join-Path $ProjectPath 'Strings\strings.psd1'
                    $hasCustomMsg = (Test-Path $strPathForCheck) -and ((Get-Content $strPathForCheck -Raw) -match "CustomMessage\s*=\s*'[^']+'")
                    if ($hasCustomMsg) { $saiwLines.Add("        CustomText = `$true") }
                }
                # BlockExecution and CloseProcessesCountdown only belong to 'with processes to close' parameter sets
                if ($Template.WelcomeDialog.BlockExecution) {
                    $saiwCloseLines.Add("        `$saiwParams.Add('BlockExecution', `$true)")
                }
                if ($Template.WelcomeDialog.CloseProcessesCountdown -gt 0) {
                    $saiwCloseLines.Add("        `$saiwParams.Add('CloseProcessesCountdown', $($Template.WelcomeDialog.CloseProcessesCountdown))")
                }
                $saiwBody       = $saiwLines -join "`n"
                $saiwCloseBody  = if ($saiwCloseLines.Count -gt 0) { "`n" + ($saiwCloseLines -join "`n") } else { '' }
                $newWelcome = "    ## Show Welcome Message, close processes if specified.`n    `$saiwParams = @{`n$saiwBody`n    }`n    if (`$adtSession.AppProcessesToClose.Count -gt 0)`n    {`n        `$saiwParams.Add('CloseProcesses', `$adtSession.AppProcessesToClose)$saiwCloseBody`n    }`n    Show-ADTInstallationWelcome @saiwParams"
            } else {
                $newWelcome = "    ## Welcome dialog disabled by org template."
            }
            # Pattern matches both original PSADT format and our re-applied format
            $scr = Set-TextBlock -Text $scr `
                -Pattern '(?s)    ## (?:Show Welcome Message[^\n]*\r?\n    \$saiwParams = @\{.*?    Show-ADTInstallationWelcome @saiwParams|Welcome dialog disabled by org template\.)' `
                -Replacement $newWelcome -Multiline

            # ── Uninstall Welcome dialog ──
            $uwSettings = if ($Template.PSObject.Properties['UninstallWelcomeDialog']) { $Template.UninstallWelcomeDialog } else { $null }
            if ($uwSettings -and $uwSettings.Enabled) {
                $uArgs = [System.Collections.Generic.List[string]]::new()
                if ($uwSettings.CloseProcessesCountdown -gt 0) { $uArgs.Add("-CloseProcessesCountdown $($uwSettings.CloseProcessesCountdown)") }
                if ($uwSettings.PersistPrompt)                 { $uArgs.Add('-PersistPrompt') }
                if ($uwSettings.BlockExecution)                { $uArgs.Add('-BlockExecution') }
                $uArgStr = if ($uArgs.Count -gt 0) { ' ' + ($uArgs -join ' ') } else { '' }
                $newUninstallWelcome = "    ## If there are processes to close, show Welcome Message before uninstalling.`n    if (`$adtSession.AppProcessesToClose.Count -gt 0)`n    {`n        Show-ADTInstallationWelcome -CloseProcesses `$adtSession.AppProcessesToClose$uArgStr`n    }"
            } else {
                $newUninstallWelcome = "    ## Uninstall welcome dialog disabled by org template."
            }
            $scr = Set-TextBlock -Text $scr `
                -Pattern '(?s)    ## (?:If there are processes to close.*?\r?\n    \}|Uninstall welcome dialog disabled by org template\.)' `
                -Replacement $newUninstallWelcome -Multiline

            # ── Progress dialog (all 3 occurrences: Install, Uninstall, Repair) ──
            if ($Template.ProgressDialog.Enabled) {
                $progressArgs = ''
                if ($Template.ProgressDialog.StatusMessage -and $Template.ProgressDialog.StatusMessage.Trim() -ne '') {
                    $progressArgs += " -StatusMessage '$($Template.ProgressDialog.StatusMessage)'"
                }
                if ($Template.ProgressDialog.StatusMessageDetail -and $Template.ProgressDialog.StatusMessageDetail.Trim() -ne '') {
                    $progressArgs += " -StatusMessageDetail '$($Template.ProgressDialog.StatusMessageDetail)'"
                }
                $newProgress = "    ## Show Progress Message (with the default message).`n    Show-ADTInstallationProgress$progressArgs"
            } else {
                $newProgress = "    ## Progress dialog disabled by org template."
            }
            $progPattern = '    ## (?:Show Progress Message[^\n]*\n    Show-ADTInstallationProgress[^\n]*|Progress dialog disabled by org template\.)'
            $offset = 0; $replaced = 0
            while ($replaced -lt 5) {
                $m = [regex]::Match($scr.Substring($offset), $progPattern)
                if (-not $m.Success) { break }
                $absIdx = $offset + $m.Index
                $scr = $scr.Substring(0, $absIdx) + $newProgress + $scr.Substring($absIdx + $m.Length)
                $offset = $absIdx + $newProgress.Length
                $replaced++
            }

            # ── Completion prompt (Post-Install) ──
            if ($Template.CompletionPrompt.Enabled) {
                $msg = $Template.CompletionPrompt.Message -replace "'", "''"
                $btn = $Template.CompletionPrompt.ButtonRightText -replace "'", "''"
                $newCompletion = "    ## Display a message at the end of the install.`n    if (!`$adtSession.UseDefaultMsi)`n    {`n        Show-ADTInstallationPrompt -Message '$msg' -ButtonRightText '$btn' -NoWait`n    }"
            } else {
                $newCompletion = "    ## Completion prompt disabled by org template."
            }
            $scr = Set-TextBlock -Text $scr `
                -Pattern '(?s)    ## (?:Display a message at the end[^\n]*\r?\n    if \(!\$adtSession\.UseDefaultMsi\)\r?\n    \{[\s\S]*?\}|Completion prompt disabled by org template\.)' `
                -Replacement $newCompletion -Multiline

            Set-Content -Path $scriptPath -Value $scr -Encoding UTF8
            Write-Host '  ✓ Invoke-AppDeployToolkit.ps1 updated' -ForegroundColor Green
        }

        Write-Host "✓ Org template '$($Template.TemplateName)' applied to project" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to apply org template: $($_.Exception.Message)"
        return $false
    }
}

function Get-AppIconFromWinget {
    param(
        [string]$ProjectPath,
        [string]$FilesPath
    )

    try {
        # Find IconUrl in any YAML file under FilesPath
        $iconUrl = $null
        $yamlFiles = Get-ChildItem -Path $FilesPath -Filter '*.yaml' -ErrorAction SilentlyContinue
        foreach ($yaml in $yamlFiles) {
            $content = Get-Content -Path $yaml.FullName -Raw -Encoding UTF8
            if ($content -match 'IconUrl:\s*(.+)') {
                $iconUrl = $matches[1].Trim()
                break
            }
        }

        if (-not $iconUrl) {
            Write-Host 'No WinGet IconUrl found in YAML — keeping default PSADT icon' -ForegroundColor DarkYellow
            return $false
        }

        Write-Host "Downloading app icon from WinGet: $iconUrl" -ForegroundColor Cyan

        # Determine save paths
        $assetsPath    = Join-Path $ProjectPath 'Assets'
        $psdtAssetsPath = Join-Path $ProjectPath 'PSAppDeployToolkit\Assets'

        foreach ($folder in @($assetsPath, $psdtAssetsPath)) {
            if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        }

        $iconDest = Join-Path $assetsPath 'AppIcon.png'

        $wr = Invoke-WebRequest -Uri $iconUrl -UseBasicParsing -ErrorAction Stop
        [System.IO.File]::WriteAllBytes($iconDest, $wr.Content)

        # Determine the icon file type from URL extension or Content-Type
        $ext = [System.IO.Path]::GetExtension($iconUrl).ToLower()
        if ($ext -notin @('.png','.ico','.jpg','.jpeg','.bmp')) { $ext = '.png' }

        # If it's an ICO, save with correct extension alongside the PNG copy
        if ($ext -eq '.ico') {
            $icoDest = Join-Path $assetsPath 'AppIcon.ico'
            [System.IO.File]::WriteAllBytes($icoDest, $wr.Content)
        }

        # Also copy to PSAppDeployToolkit\Assets so the toolkit's own default is replaced
        $psdtIconDest = Join-Path $psdtAssetsPath 'AppIcon.png'
        Copy-Item -Path $iconDest -Destination $psdtIconDest -Force

        Write-Host "✓ App icon downloaded and applied to Assets\" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Could not download app icon: $($_.Exception.Message)"
        return $false
    }
}

#endregion ── Org Template functions ─────────────────────────────────────────

# Main script execution
try {
    Write-Host "Winget App Downloader + PSADT Project Creator" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan
    
    # Check if winget is installed
    if (-not (Test-WingetInstalled)) {
        throw "Winget is not installed or not available in PATH"
    }

    # Load (or create) org template once for this run
    $script:OrgTemplate = Get-OrgTemplate

    # -Id fast path: skip search entirely
    if ($Id) {
        Write-Host "Resolving package ID: $Id" -ForegroundColor Yellow
        $showOutput = winget show --id "$Id" --exact --accept-source-agreements | Out-String
        if ($LASTEXITCODE -ne 0 -or $showOutput -notmatch 'Found') {
            Write-Host "Package ID '$Id' not found in winget. Verify the ID and try again." -ForegroundColor Red
            return
        }
        # Parse Name and Version from winget show output
        $resolvedName    = if ($showOutput -match '(?m)^Found\s+(.+?)\s+\[') { $matches[1].Trim() } else { $Id }
        $resolvedVersion = if ($showOutput -match '(?m)^\s*Version:\s*(.+)') { $matches[1].Trim() } else { 'Unknown' }
        $selectedApp = [PSCustomObject]@{
            Name    = $resolvedName
            Id      = $Id
            Version = $resolvedVersion
            Source  = 'winget'
        }
        Write-Host "`nSelected: $($selectedApp.Name) v$($selectedApp.Version)" -ForegroundColor Cyan
    }
    else {
        # Get search term if not provided
        if (-not $SearchTerm) {
            $SearchTerm = Read-Host "Enter search term for applications"
            if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
                throw "Search term is required"
            }
        }
        
        # Search for apps and exclude Microsoft Store results
        $apps = Search-WingetApps -SearchTerm $SearchTerm
        $apps = $apps | Where-Object { $_.Source -ne 'msstore' }
        
        if ($apps.Count -eq 0) {
            Write-Host "No applications found matching: $SearchTerm" -ForegroundColor Yellow
            return
        }
        
        Write-Host "Found $($apps.Count) applications" -ForegroundColor Green

        # Display apps as a numbered list with alternating row colours
        $evenColor = 'Cyan'
        $oddColor  = 'White'

        Write-Host ""
        Write-Host ("  {0,-3} {1,-28} {2,-32} {3,-10} {4}" -f "#", "Name", "Id", "Version", "Source") -ForegroundColor Gray
        Write-Host ("  " + "-" * 82) -ForegroundColor DarkGray

        for ($i = 0; $i -lt $apps.Count; $i++) {
            $color = if ($i % 2 -eq 0) { $evenColor } else { $oddColor }
            $row = "  {0,-3} {1,-28} {2,-32} {3,-10} {4}" -f ($i + 1), $apps[$i].Name, $apps[$i].Id, $apps[$i].Version, $apps[$i].Source
            Write-Host $row -ForegroundColor $color
        }
        Write-Host ""

        do {
            $rawInput = Read-Host "Select application (1-$($apps.Count))"
            $parsed   = 0
            $valid    = [int]::TryParse($rawInput.Trim(), [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $apps.Count
            if (-not $valid) {
                Write-Host "Invalid selection. Please enter a number between 1 and $($apps.Count)." -ForegroundColor Red
            }
        } while (-not $valid)

        $selectedApp = $apps[$parsed - 1]

        if (-not $selectedApp) {
            Write-Host "No application selected" -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nSelected: $($selectedApp.Name) v$($selectedApp.Version)" -ForegroundColor Cyan
    }
    
    # Get available architectures for the selected app
    $availableArchs = Get-WingetAppDetails -AppId $selectedApp.Id
    
    # Let user select architecture (single selection for project creation)
    $selectedArch = Select-Architecture -Architectures $availableArchs -AppName $selectedApp.Name
    
    if ($selectedArch -eq "all") {
        Write-Host "For project creation, please select a specific architecture." -ForegroundColor Yellow
        $selectedArch = Select-Architecture -Architectures $availableArchs -AppName $selectedApp.Name
    }
    
    # Create project name using AppName_Architecture_Version format
    $appNameClean = Sanitize-ProjectName -Name $selectedApp.Name
    $versionClean = Sanitize-ProjectName -Name $selectedApp.Version
    $archClean = Sanitize-ProjectName -Name $selectedArch
    
    $projectName = "${appNameClean}_${archClean}_${versionClean}"
    
    Write-Host "`nProject name will be: $projectName" -ForegroundColor Cyan
    
    # Create project first
    $projectCreated = Create-PSADTProject -ProjectName $projectName -ProjectPath $BasePath
    
    if ($projectCreated) {
        # Create download path within the project's Files directory
        $projectFullPath = Join-Path $BasePath $projectName
        $downloadPath = Join-Path $projectFullPath "Files"
        
        # Ensure Files directory exists
        if (!(Test-Path $downloadPath)) {
            New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
        }
        
        Write-Host "`nDownloading application to project Files directory..." -ForegroundColor Yellow
        
        # Download the selected app with selected architecture
        $downloadSuccess = Download-WingetApp -AppId $selectedApp.Id -AppName $selectedApp.Name -DownloadPath $downloadPath -Architecture $selectedArch
        
        if ($downloadSuccess) {
            Write-Host "`n✓ App downloaded successfully!" -ForegroundColor Green
            
            # Configure PSADT based on downloaded files
            Write-Host "Configuring PSADT for installer type..." -ForegroundColor Yellow
            $psadtConfigured = Configure-PSADTForInstaller -ProjectPath $projectFullPath -AppInfo $selectedApp -Architecture $selectedArch

            # Download and apply the app-specific icon from WinGet YAML
            Write-Host 'Downloading app icon from WinGet...' -ForegroundColor Cyan
            Get-AppIconFromWinget -ProjectPath $projectFullPath -FilesPath $downloadPath | Out-Null

            if ($psadtConfigured) {
                Write-Host "`n✓ SUCCESS: Project created, app downloaded, and PSADT configured!" -ForegroundColor Green
            } else {
                Write-Host "`n✓ SUCCESS: Project created and app downloaded!" -ForegroundColor Green
                Write-Warning "PSADT configuration had issues - please review manually"
            }
            
            Write-Host "Project location: $projectFullPath" -ForegroundColor Cyan
            Write-Host "Downloaded files: $downloadPath" -ForegroundColor Cyan
            
            # Generate targeted installation documentation (always runs)
            Write-Host "`nGenerating targeted installation documentation..." -ForegroundColor Yellow
            $docSuccess = New-TargetedDocumentation -ProjectPath $projectFullPath -ProjectName $projectName -AppInfo $selectedApp
            
            if ($docSuccess) {
                Write-Host "✓ Targeted documentation setup completed!" -ForegroundColor Green
                Write-Host "Use the generated Windows Sandbox configuration to document installation changes." -ForegroundColor Cyan
                
                # Wait for documentation completion and process the results
                Write-Host "`nWaiting for documentation completion..." -ForegroundColor Yellow
                $fileInfo = Get-InstallerFileInfo -FilesPath $downloadPath
                $jsonProcessed = Wait-ForDocumentationAndProcess -ProjectPath $projectFullPath -InstallerType $fileInfo.Type
                
                if ($jsonProcessed) {
                    Write-Host "✓ Documentation processing completed successfully!" -ForegroundColor Green
                } else {
                    Write-Warning "Documentation processing had issues - please review manually"
                }
            } else {
                Write-Warning "Documentation generation had issues - please review manually"
            }
            
            # Open project folder
            if (Get-Command explorer.exe -ErrorAction SilentlyContinue) {
                explorer.exe $projectFullPath
            }
        }
        else {
            Write-Warning "Project was created but download failed. You can manually add the application files to the Files directory."
        }
    }
    else {
        Write-Error "Failed to create PSADT project"
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
}