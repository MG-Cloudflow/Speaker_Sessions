#Requires -Version 5.1

<#
.SYNOPSIS
    Creates a new PSADT V4 template project using PowerShell module
.DESCRIPTION
    Uses the PSAppDeployToolkit PowerShell module to create a new V4 template
.PARAMETER ProjectName
    Name of the new PSADT project
.PARAMETER ProjectPath
    Path where the project will be created
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectName,
    
    [Parameter(Mandatory = $false)]
    [string]$ProjectPath = "C:\Github\Speaker_Sessions\wpnsummit\Win32PSADT\Demo1"
)

# Variables
$ProjectFullPath = Join-Path $ProjectPath $ProjectName

try {
    Write-Host "Creating new PSADT V4 project using PowerShell module: $ProjectName" -ForegroundColor Green
    
    # Check if PSADT module is installed
    Write-Host "Checking for PSAppDeployToolkit module..." -ForegroundColor Yellow
    $PSADTModule = Get-Module -Name PSAppDeployToolkit -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $PSADTModule) {
        Write-Host "PSAppDeployToolkit module not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name PSAppDeployToolkit -Scope CurrentUser -Force -AllowClobber
        $PSADTModule = Get-Module -Name PSAppDeployToolkit -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    }
    
    Write-Host "Using PSAppDeployToolkit version: $($PSADTModule.Version)" -ForegroundColor Green
    
    # Import the module
    Import-Module -Name PSAppDeployToolkit -Force
    
    # Create project directory
    if (!(Test-Path $ProjectPath)) {
        New-Item -Path $ProjectPath -ItemType Directory -Force | Out-Null
    }
    
    if (Test-Path $ProjectFullPath) {
        throw "Project directory already exists: $ProjectFullPath"
    }
    
    # Create new PSADT template using the module
    Write-Host "Creating PSADT V4 template..." -ForegroundColor Yellow
    New-ADTTemplate -Destination $ProjectPath -Name $ProjectName
    
    Write-Host "PSADT V4 project created successfully!" -ForegroundColor Green
    Write-Host "Project location: $ProjectFullPath" -ForegroundColor Cyan
    Write-Host "Main script: $(Join-Path $ProjectFullPath 'Deploy-Application.ps1')" -ForegroundColor Cyan
    
    # Open project folder
    if (Get-Command explorer.exe -ErrorAction SilentlyContinue) {
        explorer.exe $ProjectFullPath
    }
}
catch {
    Write-Error "Failed to create PSADT project: $($_.Exception.Message)"
}
