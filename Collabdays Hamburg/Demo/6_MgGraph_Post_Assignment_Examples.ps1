$ConfigurationProfileAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/3128dab4-2bb3-4c15-88f2-cd6a24d3572e/assignments"

# Define the configuration policy ID and the group ID to be assigned
$configurationPolicyId = "3128dab4-2bb3-4c15-88f2-cd6a24d3572e"
$groupId = "080ae94b-d551-43ea-8f95-c43e4e68feb7"  # Replace with the actual group ID

# Define the URL for the assign action
$url = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$configurationPolicyId')/assign"

# Define the request body with the group assignment
$body = @{
    assignments = @(
        @{
            target = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                groupId = $groupId
            }
        }
    )
} | ConvertTo-Json -Depth 10

# Send the POST request to assign the group to the configuration policy
$response = Invoke-MgGraphRequest -Method Post -Uri $url -Body $body -ContentType "application/json"

# Output the response
$response