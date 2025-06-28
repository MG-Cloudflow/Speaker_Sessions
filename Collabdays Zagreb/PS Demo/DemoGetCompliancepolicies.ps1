# Demo Script: Retrieve Windows Compliance Policies from Intune using Microsoft Graph API

# Step 1: Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"

# Step 2: Verify connection
$connectionStatus = Get-MgUser -UserId "4f23ee40-7c8a-4def-b4d9-489dc7de3201"
if ($connectionStatus) {
    Write-Host "Connected as: $($connectionStatus.DisplayName)"
} else {
    Write-Host "Connection failed. Ensure you have the correct permissions."
    exit
}

# Step 3: Define the Graph API URL for Windows compliance policies
$intuneCompliancePolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"

Write-Host "Fetching all Windows Compliance Policies from Microsoft Intune..."

# Step 4: Use Invoke-MgGraphRequest to make the API call
$compliancePolicies = Invoke-MgGraphRequest -Method GET -Uri $intuneCompliancePolicyUrl

# Step 5: Clean up and disconnect
Disconnect-MgGraph

Write-Host "Disconnected. Thank you!"
