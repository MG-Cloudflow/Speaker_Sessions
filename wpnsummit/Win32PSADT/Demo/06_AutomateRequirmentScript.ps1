<#
.SYNOPSIS
    Generates Intune requirement scripts based on InstallationChanges JSON data
.DESCRIPTION
    This script reads the InstallationChanges JSON file and creates PowerShell detection scripts
    that can be used as Intune Win32 app requirements to check if an application is installed.
.PARAMETER ProjectPath
    Path to the PSADT project containing the InstallationChanges JSON file
.PARAMETER OutputPath
    Optional path to save the requirement script. If not specified, saves to SupportFiles directory as RequirementScript.ps1
.PARAMETER TestMode
    Run in test mode to preview the generated requirement script without saving
.EXAMPLE
    .\Create-IntuneRequirement.ps1 -ProjectPath "C:\PSADT\MyApp" -TestMode
.EXAMPLE
    .\Create-IntuneRequirement.ps1 -ProjectPath "C:\PSADT\MyApp" -OutputPath "C:\Requirements\MyApp-Requirement.ps1"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_})]
    [string]$ProjectPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$TestMode
)

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

try {
    Write-ColorOutput "Intune Requirement Generator" "Cyan"
    Write-ColorOutput "============================" "Cyan"
    
    # Find InstallationChanges JSON file
    Write-ColorOutput "Searching for JSON files..." "White"
    $jsonFiles = @()
    
    # Search in project root
    $jsonFiles += Get-ChildItem -Path $ProjectPath -Filter "InstallationChanges*.json" -File -ErrorAction SilentlyContinue
    
    # Search in Documentation subfolder
    $docPath = Join-Path $ProjectPath "Documentation"
    if (Test-Path $docPath) {
        $jsonFiles += Get-ChildItem -Path $docPath -Filter "InstallationChanges*.json" -File -ErrorAction SilentlyContinue
    }
    
    $jsonFiles = $jsonFiles | Sort-Object LastWriteTime -Descending
    
    if ($jsonFiles.Count -eq 0) {
        throw "No InstallationChanges JSON files found in: $ProjectPath or $docPath"
    }
    
    $jsonFile = $jsonFiles[0]
    Write-ColorOutput "Using: $($jsonFile.Name)" "Green"
    
    # Parse JSON
    Write-ColorOutput "Parsing JSON data..." "White"
    $jsonContent = Get-Content -Path $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    
    # Extract application info
    $appEntries = @()
    $productCodes = @()
    
    # Handle different JSON structures
    if ($jsonContent.Values -and $jsonContent.Values.PSObject.Properties) {
        # Direct Values structure
        foreach ($prop in $jsonContent.Values.PSObject.Properties) {
            $entry = $prop.Value
            if ($entry.DisplayName) {
                $appEntries += $entry
                
                # Extract product codes from UninstallString if present
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
                
                # Extract product codes from UninstallString if present
                if ($entry.UninstallString -and $entry.UninstallString -match '\{[A-F0-9-]{36}\}') {
                    $productCodes += $matches[0]
                }
                
                # Also check if the registry path contains a GUID (product code)
                if ($regItem.Path -and $regItem.Path -match '\{[A-F0-9-]{36}\}') {
                    $productCodes += $matches[0]
                }
            }
        }
    }
    
    if ($appEntries.Count -eq 0) {
        throw "No application entries found in JSON file"
    }
    
    # Use the first/main application entry
    $mainApp = $appEntries[0]
    $appName = $mainApp.DisplayName
    $appVersion = $mainApp.DisplayVersion
    $publisher = $mainApp.Publisher
    
    Write-ColorOutput "Extracted info for: $appName" "Green"
    if ($appVersion) { Write-ColorOutput "  Version: $appVersion" "White" }
    if ($publisher) { Write-ColorOutput "  Publisher: $publisher" "White" }
    Write-ColorOutput "  Registry Entries: $($appEntries.Count)" "White"
    if ($productCodes.Count -gt 0) { Write-ColorOutput "  Product Codes: $($productCodes.Count)" "White" }
    
    # Generate requirement script
    Write-ColorOutput "Generating requirement script..." "White"
    
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

    if ($TestMode) {
        Write-ColorOutput "`nTEST MODE: Generated requirement script:" "Cyan"
        Write-Host $requirementScript -ForegroundColor Gray
        Write-ColorOutput "`nScript Logic Summary:" "Yellow"
        if ($productCodes.Count -gt 0) {
            Write-ColorOutput "  1. Checks specific MSI product codes first" "White"
        }
        Write-ColorOutput "  2. Searches registry for application by name" "White"
        if ($appVersion) {
            Write-ColorOutput "  3. Compares installed version with required version ($appVersion)" "White"
        }
        Write-ColorOutput "  4. Returns exit code 0 if requirement met, 1 if not" "White"
        return
    }
    
    # Save or output the script
    if ($OutputPath) {
        # Ensure directory exists
        $outputDir = Split-Path -Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }
        
        $requirementScript | Set-Content -Path $OutputPath -Encoding UTF8
        Write-ColorOutput "`n✓ SUCCESS: Requirement script saved to:" "Green"
        Write-ColorOutput "  $OutputPath" "White"
    } else {
        # Generate default filename in SupportFiles directory
        $supportFilesPath = Join-Path $ProjectPath "SupportFiles"
        if (-not (Test-Path $supportFilesPath)) {
            New-Item -Path $supportFilesPath -ItemType Directory -Force | Out-Null
            Write-ColorOutput "Created SupportFiles directory" "Yellow"
        }
        
        $defaultPath = Join-Path $supportFilesPath "RequirementScript.ps1"
        
        $requirementScript | Set-Content -Path $defaultPath -Encoding UTF8
        Write-ColorOutput "`n✓ SUCCESS: Requirement script saved to:" "Green"
        Write-ColorOutput "  $defaultPath" "White"
    }
    
    Write-ColorOutput "`nUsage Instructions:" "Cyan"
    Write-ColorOutput "1. Copy the generated PowerShell script content" "White"
    Write-ColorOutput "2. In Intune, go to your Win32 app > Requirements" "White"
    Write-ColorOutput "3. Add requirement rule: 'Script'" "White"
    Write-ColorOutput "4. Paste the script content" "White"
    Write-ColorOutput "5. Set 'Run script as 32-bit process': No" "White"
    Write-ColorOutput "6. Set 'Enforce script signature check': No" "White"
    
} catch {
    Write-ColorOutput "`n❌ Error: $($_.Exception.Message)" "Red"
    exit 1
}