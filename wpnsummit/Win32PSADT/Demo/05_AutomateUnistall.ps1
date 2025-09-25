<#
.SYNOPSIS
    Simple PSADT Uninstall Generator
.DESCRIPTION  
    Generates PSADT uninstall code from InstallationChanges JSON files
.PARAMETER ProjectPath
    Path to PSADT project directory
.PARAMETER TestMode
    P        $codeLines += "            } catch {"
        $codeLines += "                Write-ADTLogEntry -Message `"Registry uninstall failed: `$(`$_.Exception.Message)`" -Severity 2"
        $codeLines += "            }"iew mode only
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,
    
    [switch]$TestMode,
    [switch]$Force
)

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    switch ($Color.ToLower()) {
        'red' { Write-Host $Message -ForegroundColor Red }
        'green' { Write-Host $Message -ForegroundColor Green }
        'yellow' { Write-Host $Message -ForegroundColor Yellow }
        'cyan' { Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host $Message }
    }
}

Write-ColorOutput "PSADT Uninstall Generator" "Cyan"
Write-ColorOutput "=========================" "Cyan"

try {
    # Find JSON files
    Write-ColorOutput "Searching for JSON files..." "Yellow"
    
    $jsonFiles = @()
    $searchPaths = @(
        "$ProjectPath\*InstallationChanges*.json",
        "$ProjectPath\Documentation\*InstallationChanges*.json"
    )
    
    foreach ($path in $searchPaths) {
        $found = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
        foreach ($file in $found) {
            $jsonFiles += $file.FullName
        }
    }
    
    if ($jsonFiles.Count -eq 0) {
        throw "No JSON files found in $ProjectPath"
    }
    
    $jsonFile = $jsonFiles[0]  # Use first found file
    Write-ColorOutput "Using: $(Split-Path $jsonFile -Leaf)" "Green"
    
    # Parse JSON
    Write-ColorOutput "Parsing JSON data..." "Yellow"
    $jsonContent = Get-Content -Path $jsonFile -Raw -Encoding UTF8
    $data = $jsonContent | ConvertFrom-Json
    
    # Extract info
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
                # Also handle Properties array structure (backward compatibility)
                elseif ($regKey.Properties) {
                    foreach ($prop in $regKey.Properties) {
                        if ($prop.Name -eq "DisplayName") {
                            $appName = $prop.Value
                        }
                        if ($prop.Name -eq "UninstallString") {
                            $uninstallStrings += $prop.Value
                        }
                    }
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
    
    Write-ColorOutput "Extracted info for: $appName" "Green"
    Write-ColorOutput "  Product Codes: $($productCodes.Count)" "White"
    Write-ColorOutput "  Uninstall Strings: $($uninstallStrings.Count)" "White"
    Write-ColorOutput "  Install Paths: $($installPaths.Count)" "White"
    
    # Generate uninstall code
    Write-ColorOutput "Generating uninstall code..." "Yellow"
    
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
            $codeLines += "                `$exitCode = Start-ADTMsiProcess -Action 'Uninstall' -Path '$pc' -Parameters '/quiet /norestart'"
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
    
    # Process registry uninstall strings and determine installer type here
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
                    $codeLines += "                `$exitCode = Start-ADTMsiProcess -Action 'Uninstall' -Path '$productCode' -Parameters '/quiet /norestart'"
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
                $codeLines += "                `$exitCode = Start-ADTProcess -Path '$exePath' -Parameters '$exeParams' -Wait -PassThru"
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
    
    if ($TestMode) {
        Write-ColorOutput "TEST MODE: Generated code:" "Cyan"
        Write-Host $uninstallCode -ForegroundColor Gray
        return
    }
    
    # Update PSADT file
    $scriptPath = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
    if (-not (Test-Path $scriptPath)) {
        throw "Invoke-AppDeployToolkit.ps1 not found: $scriptPath"
    }
    
    # Backup
    $backupPath = "$scriptPath.backup.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item -Path $scriptPath -Destination $backupPath -Force
    Write-ColorOutput "✓ Created backup" "Green"
    
    # Confirm unless Force
    if (-not $Force) {
        Write-ColorOutput "`nReady to update PSADT for: $appName" "Cyan"
        $confirm = Read-Host "Proceed? (y/n)"
        if ($confirm -notin @('y', 'Y')) {
            Write-ColorOutput "Cancelled" "Yellow"
            return
        }
    }
    
    # Read and update
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
    
    Write-ColorOutput "`n✓ SUCCESS: Updated PSADT for $appName" "Green"
    Write-ColorOutput "Methods: $(if($productCodes.Count -gt 0){'MSI '})$(if($uninstallStrings.Count -gt 0){'Registry '})Cleanup" "White"
    
} catch {
    Write-ColorOutput "`n❌ Error: $($_.Exception.Message)" "Red"
    exit 1
}