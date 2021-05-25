#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

# Reference: https://docs.microsoft.com/en-us/graph/api/device-update?view=graph-rest-1.0&tabs=http
# Create an App Registration 
# Assign Delegated permiossions: Directory.ReadWrite.All or Directory.AccessAsUser.All
# Update $vars variable below: Client Secret, ClientID and TenantID

# This script will dismiss all risky users.

# Variable: Store variables used throughout this script
$vars = @{
    # Used to generate an Access Token to query Microsoft Graph API.
    Token = @{
        ClientSecret = ""
        ClientID = ""
        TenantID = ""
        AccessToken = ""
    }

    DeviceObjectID = ""
    
    ScriptStartTime = Get-Date
}

# Function: Create Access Token for Microsoft Graph API
function Get-AppToken{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $clientId,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $clientSecret,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $tenantId

    )

    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth 2.0 Token
    $tokenRequestArgs = @{
        Method = "Post"
        Uri = $uri 
        ContentType = "application/x-www-form-urlencoded" 
        Body = $body 
        UseBasicParsing = $true
    }
    $tokenRequest = Invoke-WebRequest @tokenRequestArgs

    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token

    return $token
}

function Enable-Device{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [string] $ObjectID
    )

    $uri = "https://graph.microsoft.com/beta/devices/$ObjectID"
    $method = "PATCh"

    $body = @{
        "accountEnabled" = $true
    } | convertTo-Json
    

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -body $body -ErrorAction Stop
    
    return $Result
}

# Get our Access Token for Microsoft Graph API
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

Enable-Device -Token $vars.Token.AccessToken -ObjectID $vars.DeviceObjectID