# Fetch all configuration policies with expanded assignments
$ConfigurationProfilesAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments"
# This command fetches all configuration policies and expands the `assignments` property.
# The `assignments` property will include detailed information about the assignments for each configuration policy.
# Result: The response will contain all configuration policies with their assignments expanded.

# Fetch all configuration policies with expanded assignments and selected properties
$ConfigurationProfilesAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments&`$select=id,assignments,name"
# This command fetches all configuration policies, expands the `assignments` property, and selects only the `id` and `assignments` properties.
# The `select` query parameter is used to limit the properties returned in the response.
# Result: The response will contain all configuration policies with only the `id` and `assignments` properties expanded.

# Fetch a specific configuration policy with expanded assignments
$ConfigurationProfileAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/3128dab4-2bb3-4c15-88f2-cd6a24d3572e/assignments"
#Because microsoft graph does not support the expand parameter for a specific configuration policy, we use the assignments endpoint directly.
$ConfigurationProfileAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/3128dab4-2bb3-4c15-88f2-cd6a24d3572e?`$expand=assignments"

# This command fetches a specific configuration policy (identified by its ID) and expands the `assignments` property.
# The ID `2202d608-5d9d-4a59-87ad-8c014b835b6f` is used to specify the configuration policy.
# Result: The response will contain the specified configuration policy with its assignments expanded.

# Fetch a specific mobileAppspolicy with expanded assignments
$MobileAppsAssignments = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/2d9dd897-d95b-469b-befc-da19746062b0?`$expand=assignments"
# This command fetches a specific configuration policy (identified by its ID) and expands the `assignments` property.
# The ID `2d9dd897-d95b-469b-befc-da19746062b0` is used to specify the configuration policy.
# Result: The response will contain the specified configuration policy with its assignments expanded.

# Fetch a specific configuration policy with expanded assignments and settings
$ConfigurationProfileAssignmentsAndSettings = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/3128dab4-2bb3-4c15-88f2-cd6a24d3572e/?`$expand=assignments,settings"

# This command fetches a specific configuration policy (identified by its ID) and expands both the `assignments` and `settings` properties.
# The ID `2202d608-5d9d-4a59-87ad-8c014b835b6f` is used to specify the configuration policy.
# Result: The response will contain the specified configuration policy with its assignments and settings expanded.
