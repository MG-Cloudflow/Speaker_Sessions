# Define the configuration policy ID and the group ID to be assigned
$configurationPolicyId = "3128dab4-2bb3-4c15-88f2-cd6a24d3572e"
$groupId = "8ac563f6-6095-4ec7-9e0d-b0fa93b49313"  # Replace with the actual group ID

# Define the URL to fetch the existing assignments
$fetchUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$configurationPolicyId')/assignments"

# Fetch the existing assignments
$existingAssignmentsResponse = Invoke-MgGraphRequest -Method Get -Uri $fetchUrl
$existingAssignments = $existingAssignmentsResponse.value

# Define the new assignment
$newAssignment = @{
    target = @{
        "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
        groupId = $groupId
    }
}

# Append the new assignment to the existing assignments
$updatedAssignments = $existingAssignments + $newAssignment

# Define the URL for the assign action
$assignUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$configurationPolicyId')/assign"

# Define the request body with the updated assignments
$body = @{
    assignments = $updatedAssignments
} | ConvertTo-Json -Depth 10

# Send the POST request to assign the group to the configuration policy
$response = Invoke-MgGraphRequest -Method Post -Uri $assignUrl -Body $body -ContentType "application/json"

# Output the response
$response