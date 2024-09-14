# Demo Script: Create a Windows Compliance Policy using Microsoft Graph API

# Step 1: Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All"

# Step 2: Verify connection
$connectionStatus = Get-MgUser -UserId "4f23ee40-7c8a-4def-b4d9-489dc7de3201"
if ($connectionStatus) {
    Write-Host "Connected as: $($connectionStatus.DisplayName)"
} else {
    Write-Host "Connection failed. Ensure you have the correct permissions."
    exit
}

# Step 3: Define the Graph API URL for Windows compliance policies
$createCompliancePolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"

# Step 4: Define the compliance policy payload

$body = @"
{
    "id": "00000000-0000-0000-0000-000000000000",
    "displayName": "Zagreb CollabDays Compliance Policy Demo",
    "roleScopeTagIds": [
        "0"
    ],
    "@odata.type": "#microsoft.graph.windows10CompliancePolicy",
    "scheduledActionsForRule": [
        {
            "ruleName": "PasswordRequired",
            "scheduledActionConfigurations": [
                {
                    "actionType": "block",
                    "gracePeriodHours": 48,
                    "notificationTemplateId": "",
                    "notificationMessageCCList": []
                }
            ]
        }
    ],
    "deviceThreatProtectionRequiredSecurityLevel": "unavailable",
    "passwordRequiredType": "deviceDefault",
    "osMinimumVersion": "$($osVersion.qualityUpdateVersion)",
    "deviceThreatProtectionEnabled": false
}
"@

# Step 5: Send the POST request to create the compliance policy
$newCompliancePolicy = Invoke-MgGraphRequest -Method POST -Uri $createCompliancePolicyUrl -Body $body

$intuneCompliancePolicyUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($newCompliancePolicy.id)')"

$CheckNewPolciy = Invoke-MgGraphRequest -Method GET -Uri $intuneCompliancePolicyUrl

# Step 7: Clean up and disconnect
Disconnect-MgGraph

