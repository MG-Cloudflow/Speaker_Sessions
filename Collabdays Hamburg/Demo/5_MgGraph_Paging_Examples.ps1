# Function to fetch data from Microsoft Graph API with pagination handling
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
$ConfigurationProfiles = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
# This command sends a single GET request to the specified URL and stores the response in $ConfigurationProfiles.
# It does not handle pagination, so it will only fetch the first page of results.
# If the response contains multiple pages of data, only the first page will be included in $ConfigurationProfiles.
# This may result in incomplete data if the total number of configuration profiles exceeds the page size limit.

# Fetch configuration profiles with pagination handling
$ConfigurationProfilesWithPaging = Get-GraphData -Url "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
# This command calls the Get-GraphData function with the same URL.
# The function handles pagination, fetching all pages of results and storing the accumulated data in $ConfigurationProfileswithpaging.
# It follows the @odata.nextLink links to retrieve all available data across multiple pages.
# This ensures that $ConfigurationProfileswithpaging contains the complete set of configuration profiles, regardless of the total number.

# Fetch all configuration policies with expanded assignments and settings
$ConfigurationProfilesAssignmentsAndSettings = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments,settings"
# This command fetches all configuration policies and expands both the `assignments` and `settings` properties.
# The `settings` property will include detailed information about the settings for each configuration policy.
# Result: TThis may result in incomplete data if the total number of configuration profiles exceeds the page size limit.


# Fetch all configuration policies with expanded assignments and settings
$ConfigurationProfilesAssignmentsAndSettings = Get-GraphData  -Url "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments,settings"
# This command fetches all configuration policies and expands both the `assignments` and `settings` properties.
# The `settings` property will include detailed information about the settings for each configuration policy.
# Result: The response will contain all configuration policies with their assignments and settings expanded.


Invoke-MgGraphRequest -Method Delete -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/236d12f5-5c84-495a-bbec-c48f1f3a5dfb"