# Define the path for the sandbox configuration file and local storage for installers
$sandboxLocalFolder = $PSScriptRoot
if (-not $sandboxLocalFolder) {
    $sandboxLocalFolder = Get-Location
}
# Determine the root path dynamically
$sandboxLocalFolder = $PSScriptRoot
if (-not $sandboxLocalFolder) {
    $sandboxLocalFolder = Get-Location
}

# Define the Sandbox folder path
$sandboxFolder = "$sandboxLocalFolder\Sandbox"

# Check if the folder exists, if not, create it
if (-not (Test-Path -Path $sandboxFolder)) {
    New-Item -ItemType Directory -Path $sandboxFolder -Force | Out-Null
    Write-Host "Sandbox folder created: $sandboxFolder"
} else {
    Write-Host "Sandbox folder already exists: $sandboxFolder"
}

# Define the SandboxConfig file path
$sandboxConfigFile = "$sandboxFolder\SandboxConfig.wsb"

Write-Host "SandboxConfig file path set to: $sandboxConfigFile"


# Create the sandbox configuration file content
$sandboxConfigContent = @"
<Configuration>
    <VGpu>Disable</VGpu>
    <Networking>Enable</Networking>
    <MappedFolders>
        <MappedFolder>
            <HostFolder>$sandboxLocalFolder</HostFolder>
            <SandboxFolder>C:\PSADT</SandboxFolder>
            <ReadOnly>false</ReadOnly>
        </MappedFolder>
    </MappedFolders>
    <LogonCommand>
        <Command>powershell.exe -NoExit -ExecutionPolicy Bypass -Command &quot;&amp; {C:\PSADT\Invoke-AppDeployToolkit.ps1 }&quot;</Command>
    </LogonCommand>
</Configuration>
"@

# Write the sandbox configuration to the file
$sandboxConfigContent | Set-Content -Path $sandboxConfigFile -Encoding UTF8

# Run the sandbox
Start-Process -FilePath "WindowsSandbox.exe" -ArgumentList $sandboxConfigFile
