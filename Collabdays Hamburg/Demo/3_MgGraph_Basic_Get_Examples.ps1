# Example 1: Fetch all Intune devices
# This example connects to Microsoft Graph with the required scope to read all managed devices in Intune.
# It then fetches all Intune devices using Invoke-MgGraphRequest and outputs the list of devices.
# User Experience: The user will be prompted to sign in and consent to the specified scopes.
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

# Fetch all Intune devices using Invoke-MgGraphRequest
$intuneDevices = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"

# Output the list of devices
$intuneDevices.value | ForEach-Object {
    Write-Host "Device ID: $($_.id)"
    Write-Host "Device Name: $($_.deviceName)"
    Write-Host "Operating System: $($_.operatingSystem)"
    Write-Host "Compliance State: $($_.complianceState)"
    Write-Host "-----------------------------"
}

# Example 2: Fetch all Windows Intune devices
# This example connects to Microsoft Graph with the required scope to read all managed devices in Intune.
# It then fetches only the Windows Intune devices using Invoke-MgGraphRequest with a filter applied in the URL.
# User Experience: The user will be prompted to sign in and consent to the specified scopes.
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

# Fetch all Windows Intune devices using Invoke-MgGraphRequest with a filter
$intuneWindowsDevices = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'"

# Output the list of Windows devices
$intuneWindowsDevices.value | ForEach-Object {
    Write-Host "Device ID: $($_.id)"
    Write-Host "Device Name: $($_.deviceName)"
    Write-Host "Operating System: $($_.operatingSystem)"
    Write-Host "Compliance State: $($_.complianceState)"
    Write-Host "-----------------------------"
}

# Example 3: Fetch all Windows Intune devices with selected properties
# This example connects to Microsoft Graph with the required scope to read all managed devices in Intune.
# It then fetches only the Windows Intune devices using Invoke-MgGraphRequest with a filter applied in the URL and selects specific properties.
# User Experience: The user will be prompted to sign in and consent to the specified scopes.
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

# Fetch all Windows Intune devices using Invoke-MgGraphRequest with a filter and select specific properties
$intuneWindowsDevices = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=id,deviceName,operatingSystem,complianceState"

# Output the list of Windows devices
$intuneWindowsDevices.value | ForEach-Object {
    Write-Host "Device ID: $($_.id)"
    Write-Host "Device Name: $($_.deviceName)"
    Write-Host "Operating System: $($_.operatingSystem)"
    Write-Host "Compliance State: $($_.complianceState)"
    Write-Host "-----------------------------"
}