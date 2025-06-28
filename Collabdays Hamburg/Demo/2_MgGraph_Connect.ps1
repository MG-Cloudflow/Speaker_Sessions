# Method 1: Connect using specific scopes
# This method connects to Microsoft Graph with the specified scopes.
# User Experience: The user will be prompted to sign in and consent to the specified scopes.

Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

# Method 2: Connect using specific scopes with device authentication
# This method connects to Microsoft Graph with the specified scopes and uses device authentication.
# User Experience: The user will be provided with a device code and a URL to enter the code. This is useful for scenarios where the user cannot directly interact with the authentication prompt.
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All" -UseDeviceAuthentication

# Method 3: Connect using client ID and tenant ID
# This method connects to Microsoft Graph using a client ID and tenant ID.
# User Experience: The user must provide the client ID and tenant ID of the application registered in Azure AD. The user will be prompted to sign in.
Connect-MgGraph -ClientId "<YOUR_NEW_APP_ID>" -TenantId "<YOUR_TENANT_ID>"

# Method 4: Connect using app-based authentication
# This method connects to Microsoft Graph using app-based authentication with client credentials.
# User Experience: The user must provide the application ID, application secret, and tenant ID. This method is useful for service-to-service calls where no user interaction is required.
function Connect-MgGraphEntraApp {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AppId,

        [Parameter(Mandatory=$true)]
        [string]$AppSecret,

        [Parameter(Mandatory=$true)]
        [string]$Tenant
    )
    
    # region app secret
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $AppId
        client_secret = $AppSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body $body
    $accessToken = $response.access_token
    $version = (Get-Module microsoft.graph.authentication | Select-Object -ExpandProperty Version).Major

    if ($version -eq 2) {
        $accesstokenfinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
    } else {
        Select-MgProfile -Name Beta
        $accesstokenfinal = $accessToken
    }
    try {
        Connect-MgGraph -AccessToken $accesstokenfinal
        Write-Host "Connected to tenant $Tenant using app-based authentication"
    }
    catch {
        Write-Warning "Error connecting to tenant $Tenant using app-based authentication, exiting..."
        return
    }
}

# Example usage of Connect-MgGraphEntraApp function
Connect-MgGraphEntraApp -AppId "<YOUR_NEW_APP_ID>" -AppSecret "<YOUR_APP_SECRET>" -Tenant "<YOUR_TENANT_ID>"

Disconnect-MgGraph