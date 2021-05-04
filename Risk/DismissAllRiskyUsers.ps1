#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

# Reference: https://docs.microsoft.com/en-us/graph/api/riskyuser-dismiss?view=graph-rest-1.0
# Create an App Registration 
# Assign Application permiossions: IdentityRiskyUser.ReadWrite.All
# Create Application secret
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

# Function: Get all risky users using Microsoft Graph API
function Get-RiskyUsers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token
    )

    $uri = "https://graph.microsoft.com/beta/identityProtection/riskyUsers"
    $method = "GET"

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

    $RiskyUsers = $Result.value

    # Loop over the pages to get all results
    $ResultNextLink = $Result."@odata.nextLink"
    while($ResultNextLink -ne $null){
        $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -ErrorAction Stop) 
        $ResultNextLink = $Result."@odata.nextLink"
        $RiskyUsers += $Result.value
    }

    return $RiskyUsers
}

# Function: Dismiss risky users using Microsoft Graph API
function Dismiss-RiskyUsers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [system.array] $UserIds
    )

    $uri = "https://graph.microsoft.com/beta/identityProtection/riskyUsers/dismiss"
    $method = "POST"

    $body = @{
        "userIds" = @(
          $userIds
        )
      } | convertTo-Json

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -body $body -ErrorAction Stop

    return $Result
}

# Get our Access Token for Microsoft Graph API
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

# Get all risky users from Microsoft Graph API
$AllRiskyUsers = Get-RiskyUsers -Token $vars.Token.AccessToken

# Create a list of risky users whose riskState is not 'dismissed'
$riskyUsersToDissmis = @()
foreach($riskyUser in $AllRiskyUsers){
    if($riskyUser.riskState -ne 'dismissed'){
        $riskyUsersToDissmis+= $riskyUser.id
    }
}

# Dismiss risky users
Dismiss-RiskyUsers -Token $vars.Token.AccessToken -UserIds $riskyUsersToDissmis