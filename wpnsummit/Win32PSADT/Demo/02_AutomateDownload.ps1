#Requires -Version 5.1

<#
.SYNOPSIS
    Search and download Winget applications
.DESCRIPTION
    Searches for Winget applications, displays results in grid view for selection, and downloads to specified folder
.PARAMETER SearchTerm
    Term to search for in Winget repository
.PARAMETER DownloadPath
    Path where downloaded applications will be stored
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SearchTerm,
    
    [Parameter(Mandatory = $false)]
    [string]$DownloadPath
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
    
    # Find the separator line to know where data starts
    $dataStartIndex = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^-+\s+-+\s+-+") {
            $dataStartIndex = $i + 1
            break
        }
    }
    
    # Parse each app line
    for ($i = $dataStartIndex; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if ($line -and $line -notmatch "^\d+\s+matches\s+found" -and $line -notmatch "^More\s+than") {
            # Split by multiple spaces to separate columns
            $parts = $line -split '\s{2,}'
            if ($parts.Count -ge 3) {
                $apps += [PSCustomObject]@{
                    Name = $parts[0].Trim()
                    Id = $parts[1].Trim()
                    Version = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "" }
                    Source = if ($parts.Count -gt 3) { $parts[3].Trim() } else { "winget" }
                }
            }
        }
    }
    
    return $apps
}

function Select-DownloadFolder {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select folder to download applications"
    $folderBrowser.ShowNewFolderButton = $true
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
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
                $arch = $matches[1].Trim()
                if ($arch -and $arch -ne "unknown" -and $architectures -notcontains $arch) {
                    $architectures += $arch
                }
            }
        }
        
        # If no architectures found in show command, try to get from manifest
        if ($architectures.Count -eq 0) {
            $architectures = @("x86", "x64", "arm64", "neutral")
        }
        
        return $architectures
    }
    catch {
        Write-Warning "Could not get architecture details for $AppId"
        return @("x86", "x64", "arm64")
    }
}

function Select-Architecture {
    param(
        [string[]]$Architectures,
        [string]$AppName
    )
    
    if ($Architectures.Count -eq 1) {
        return $Architectures[0]
    }
    
    Write-Host "`nAvailable architectures for $AppName :" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Architectures.Count; $i++) {
        Write-Host "$($i + 1). $($Architectures[$i])" -ForegroundColor White
    }
    
    do {
        $selection = Read-Host "Select architecture (1-$($Architectures.Count)) or 'all' for all architectures"
        
        if ($selection.ToLower() -eq "all") {
            return "all"
        }
        
        if ($selection -match '^\d+$') {
            $index = [int]$selection - 1
            if ($index -ge 0 -and $index -lt $Architectures.Count) {
                return $Architectures[$index]
            }
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
    
    Write-Host "Downloading $AppName ($AppId)..." -ForegroundColor Green
    
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

# Main script execution
try {
    Write-Host "Winget App Downloader" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    
    # Check if winget is installed
    if (-not (Test-WingetInstalled)) {
        throw "Winget is not installed or not available in PATH"
    }
    
    # Get search term if not provided
    if (-not $SearchTerm) {
        $SearchTerm = Read-Host "Enter search term for applications"
        if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
            throw "Search term is required"
        }
    }
    
    # Search for apps
    $apps = Search-WingetApps -SearchTerm $SearchTerm
    
    if ($apps.Count -eq 0) {
        Write-Host "No applications found matching: $SearchTerm" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($apps.Count) applications" -ForegroundColor Green
    
    # Display apps in grid view for selection
    $selectedApps = $apps | Out-GridView -Title "Select applications to download" -PassThru
    
    if (-not $selectedApps -or $selectedApps.Count -eq 0) {
        Write-Host "No applications selected" -ForegroundColor Yellow
        return
    }
    
    # Get download folder if not provided
    if (-not $DownloadPath) {
        $DownloadPath = Select-DownloadFolder
        if (-not $DownloadPath) {
            Write-Host "No download folder selected" -ForegroundColor Yellow
            return
        }
    }
    
    # Ensure download path exists
    if (!(Test-Path $DownloadPath)) {
        New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
    }
    
    Write-Host "Downloading $($selectedApps.Count) application(s) to: $DownloadPath" -ForegroundColor Cyan
    
    # Download each selected app with architecture selection
    $successCount = 0
    foreach ($app in $selectedApps) {
        Write-Host "`n--- Processing: $($app.Name) ---" -ForegroundColor Cyan
        
        # Get available architectures for this app
        $availableArchs = Get-WingetAppDetails -AppId $app.Id
        
        if ($availableArchs.Count -gt 0) {
            # Let user select architecture
            $selectedArch = Select-Architecture -Architectures $availableArchs -AppName $app.Name
            
            if ($selectedArch -eq "all") {
                # Download all architectures
                foreach ($arch in $availableArchs) {
                    Write-Host "Downloading $($app.Name) - $arch architecture..." -ForegroundColor Yellow
                    if (Download-WingetApp -AppId $app.Id -AppName "$($app.Name)_$arch" -DownloadPath $DownloadPath -Architecture $arch) {
                        $successCount++
                    }
                }
            }
            else {
                # Download selected architecture
                if (Download-WingetApp -AppId $app.Id -AppName $app.Name -DownloadPath $DownloadPath -Architecture $selectedArch) {
                    $successCount++
                }
            }
        }
        else {
            # Download without architecture specification
            if (Download-WingetApp -AppId $app.Id -AppName $app.Name -DownloadPath $DownloadPath) {
                $successCount++
            }
        }
    }
    
    Write-Host "`nDownload Summary:" -ForegroundColor Cyan
    Write-Host "Successfully downloaded: $successCount/$($selectedApps.Count) applications" -ForegroundColor Green
    Write-Host "Download location: $DownloadPath" -ForegroundColor Cyan
    
    # Open download folder
    if (Get-Command explorer.exe -ErrorAction SilentlyContinue) {
        explorer.exe $DownloadPath
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
}

