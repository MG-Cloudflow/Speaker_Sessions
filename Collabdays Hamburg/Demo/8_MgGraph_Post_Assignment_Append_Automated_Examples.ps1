# Define the configuration policy ID and the group ID to be assigned
#$configurationPolicyId = "3128dab4-2bb3-4c15-88f2-cd6a24d3572e"
$groupId = "080ae94b-d551-43ea-8f95-c43e4e68feb7"  # Replace with the actual group ID

function Get-GraphData {
    param (
        [Parameter(Mandatory=$true)]
        [string] $url  # URL to fetch data from
    )
    Write-Host "Starting Get-GraphData for $url"
    try {
        $results = @()  # Initialize an empty array to store results
        do {
            Write-Host "Fetching data from $url"
            $response = Invoke-MgGraphRequest -Uri $url -Method GET  # Send GET request to the URL
            if ($response.'@odata.nextLink' -ne $null) {  # Check if there is a next page
                $url = $response.'@odata.nextLink'  # Update URL to the next page URL
                Write-Host "Next page URL: $url"
                $results += $response.value  # Append current page's data to results
            } else {
                $results += $response.Value  # Append final page's data to results
                Write-Host "Successfully fetched all data"
                return $results  # Return the accumulated results
            }
        } while ($response.'@odata.nextLink')  # Continue loop if there is a next page
    } catch {
        $errorMessage = "Failed to get data from Graph API: $($_.Exception.Message)"
        Write-Error $errorMessage  # Log any errors that occur
    }
}

# Fetch configuration profiles without pagination handling
$ConfigurationProfiles = Get-GraphData -Url "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=startswith(name,'V5-IMS')&`$select=id,name"


$ConfigurationProfiles | ForEach-Object{
    # Define the URL to fetch the existing assignments
    $configurationPolicyId = $_.id
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
    $alreadyAssigned = $existingAssignments | Where-Object { $_.target.groupId -eq $groupId }

    if (-not $alreadyAssigned) {
        # Only add the new assignment if not already present
        $updatedAssignments = $existingAssignments + $newAssignment
    } else {
        # No change needed if already assigned
        $updatedAssignments = $existingAssignments
    }

    # Define the URL for the assign action
    $assignUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$configurationPolicyId')/assign"

    # Define the request body with the updated assignments
    $body = @{
        assignments = $updatedAssignments
    } | ConvertTo-Json -Depth 10

    # Send the POST request to assign the group to the configuration policy
    Invoke-MgGraphRequest -Method Post -Uri $assignUrl -Body $body -ContentType "application/json"
}
