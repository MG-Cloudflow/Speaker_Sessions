#Requires -Version 5.1

<#
.SYNOPSIS
    Documents installation changes using targeted monitoring of key system areas
.DESCRIPTION
    Uses focused monitoring of specific directories and registry locations where
    software installations typically make changes. Much faster than full system
    scans while maintaining high accuracy.
    
    Monitors:
    - Program Files directories
    - ProgramData and AppData folders
    - Key registry locations (Uninstall, App Paths, etc.)
    - Services and installed programs
    
.PARAMETER ProjectPath
    Path to the PSADT project folder containing Invoke-AppDeployToolkit.ps1
.PARAMETER OutputPath
    Path where documentation will be saved (defaults to project folder)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ProjectPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

function Select-ProjectFolder {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select PSADT project folder to document"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Get the project folder path
if (-not $ProjectPath) {
    Write-Host "Please select the PSADT project folder to document..." -ForegroundColor Yellow
    $ProjectPath = Select-ProjectFolder
    
    if (-not $ProjectPath) {
        Write-Host "No folder selected. Exiting." -ForegroundColor Red
        return
    }
}

# Validate the selected path
if (-not (Test-Path $ProjectPath)) {
    Write-Error "Project path does not exist: $ProjectPath"
    return
}

# Check if Invoke-AppDeployToolkit.ps1 exists in the selected folder
$deployScript = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
if (-not (Test-Path $deployScript)) {
    Write-Error "Invoke-AppDeployToolkit.ps1 not found in: $ProjectPath"
    return
}

# Set output path
if (-not $OutputPath) {
    $OutputPath = $ProjectPath
}

# Get project name for display purposes
$projectName = Split-Path $ProjectPath -Leaf
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Define the Sandbox folder path (create in script directory)
$scriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
$sandboxFolder = Join-Path $scriptRoot "Sandbox"

# Check if the folder exists, if not, create it
if (-not (Test-Path -Path $sandboxFolder)) {
    New-Item -ItemType Directory -Path $sandboxFolder -Force | Out-Null
    Write-Host "Sandbox folder created: $sandboxFolder" -ForegroundColor Green
} else {
    Write-Host "Sandbox folder already exists: $sandboxFolder" -ForegroundColor Gray
}

# Define the SandboxConfig file path
$sandboxConfigFile = Join-Path $sandboxFolder "Targeted_Documentation_$projectName`_$timestamp.wsb"

Write-Host "Creating targeted documentation sandbox for project: $projectName" -ForegroundColor Cyan
Write-Host "SandboxConfig file: $sandboxConfigFile" -ForegroundColor Gray
Write-Host "Project folder: $ProjectPath" -ForegroundColor Gray

# Create PowerShell targeted documentation script content
$documentationScript = @'
# PSADT Installation Documentation Script with Targeted Monitoring
$ErrorActionPreference = 'Continue'
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

# Function to export registry data in custom format (no table truncation)
function Export-RegistryDataSafely {
    param([string]$FilePath, [object]$Data, [string]$Description)
    try {
        $itemCount = if ($Data -is [array]) { $Data.Count } else { 1 }
        Write-Log "Attempting to export $Description ($itemCount items) to: $FilePath" "INFO"
        
        # Create custom formatted output to avoid PowerShell table truncation
        $output = @()
        $output += "================================================================================"
        $output += "$Description"
        $output += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $output += "Total Entries: $itemCount"
        $output += "================================================================================"
        $output += ""
        
        foreach ($item in $Data) {
            $output += "Registry Entry $($Data.IndexOf($item) + 1):"
            $output += "  Full Path: $($item.FullPath)"
            $output += "  Key Name: $($item.KeyName)"
            $output += "  Properties:$($item.Properties)"
            $output += "  Summary: $($item.Summary)"
            $output += ""
        }
        
        $output | Set-Content -Path $FilePath -Encoding UTF8 -Force
        
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
        # Compare using FullPath property (not Path)
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
    
    # Registry data kept in memory for JSON export only
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
    
    # Only export the main InstallationChanges JSON - no summary needed
    
} catch {
    Write-Host "Warning: Could not export JSON documentation" -ForegroundColor Yellow
    Write-Log "Warning: Could not export JSON documentation: $($_.Exception.Message)" "WARNING"
}

# No text summary file - only JSON and log files

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
Write-Host "• File system changes (new, modified files)" -ForegroundColor White
Write-Host "• Registry changes (new entries)" -ForegroundColor White
Write-Host "• Service and program changes" -ForegroundColor White
Write-Host "• Detailed before/after comparisons" -ForegroundColor White
Write-Host "• JSON format for automation and diff comparison" -ForegroundColor White
Write-Host "• Complete execution log" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Copy the Documentation folder to your host system" -ForegroundColor Green
Write-Host "2. Review NEW_* files to see exactly what was installed" -ForegroundColor Green
Write-Host "3. Use MODIFIED_* files to see what was changed" -ForegroundColor Green
Write-Host "4. Compare PRE_/POST_ files for detailed analysis" -ForegroundColor Green

Write-Host "`nPress any key to open documentation folder..."
Read-Host
Start-Process -FilePath "explorer.exe" -ArgumentList $outputPath
'@

# Replace placeholders with actual values
$documentationScript = $documentationScript -replace '__PROJECTNAME__', $projectName
$documentationScript = $documentationScript -replace '__TIMESTAMP__', $timestamp

# Save the documentation script inside the PSADT project's SupportFiles folder
$supportFilesFolder = Join-Path $ProjectPath "SupportFiles"
$docScriptPath = Join-Path $supportFilesFolder "TargetedDocumentationScript.ps1"

# Ensure SupportFiles folder exists
if (-not (Test-Path $supportFilesFolder)) {
    New-Item -ItemType Directory -Path $supportFilesFolder -Force | Out-Null
    Write-Host "Created SupportFiles folder: $supportFilesFolder" -ForegroundColor Green
}

$documentationScript | Set-Content -Path $docScriptPath -Encoding UTF8
Write-Host "Targeted documentation script created: $docScriptPath" -ForegroundColor Green

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
Write-Host "Targeted documentation sandbox configuration created!" -ForegroundColor Green

# Check if Windows Sandbox is available
try {
    $sandboxFeature = Get-WindowsOptionalFeature -Online -FeatureName "Containers-DisposableClientVM"
    if ($sandboxFeature.State -ne "Enabled") {
        Write-Warning "Windows Sandbox feature is not enabled. Please enable it in Windows Features."
        Write-Host "To enable: Control Panel > Programs > Turn Windows features on or off > Windows Sandbox" -ForegroundColor Yellow
        return
    }
} catch {
    Write-Warning "Unable to check Windows Sandbox feature status."
}

# Run the sandbox
Write-Host "`nStarting Targeted Documentation Sandbox..." -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "This sandbox will:" -ForegroundColor White
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
    Write-Host "Targeted Documentation Sandbox launched successfully!" -ForegroundColor Green
    Write-Host "`nThe targeted documentation will be available in the sandbox at:" -ForegroundColor Cyan
    Write-Host "C:\PSADT\Documentation" -ForegroundColor White
} catch {
    Write-Error "Failed to start Windows Sandbox: $($_.Exception.Message)"
    Write-Host "Make sure Windows Sandbox is installed and enabled." -ForegroundColor Yellow
}