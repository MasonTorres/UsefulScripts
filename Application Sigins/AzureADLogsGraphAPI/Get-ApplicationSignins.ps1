#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

# Create a new Application in Azure AD
# Assign Microsoft Graph Application Permission: AuditLog.Read.All
# Create an Application Secert
# Fill in the veriables below. (leave Access Token blank)

$vars = @{
    # Used to generate an Access Token to query Microsoft Graph API
    Token = @{
        ClientSecret = ""
        ClientID = ""
        TenantID = ""
        AccessToken = ""
    }
    # CSV Export location for the audit
    Export = @{
        AppCSVLoc = "Sign-InEvents-Applications.csv"
        SPCSVLoc = "Sign-InEvents-ServicePrincipals.csv"
    }
    ScriptStartTime = Get-Date
}

#Functions
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

function Get-Applications{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token
    )

    $uri = "https://graph.microsoft.com/beta/applications"
    $method = "GET"

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    $Applications = $Result.value

    $ResultNextLink = $Result."@odata.nextLink"
    while($ResultNextLink -ne $null){
        $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -ErrorAction Stop) 
        $ResultNextLink = $Result."@odata.nextLink"
        $Applications += $Result.value
    }

    return $Applications
}

function Get-ServicePrincipals{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token
    )

    $uri = "https://graph.microsoft.com/beta/serviceprincipals"
    $method = "GET"

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

    $ServicePrincipals = $Result.value

    $ResultNextLink = $Result."@odata.nextLink"
    while($ResultNextLink -ne $null){
        $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -ErrorAction Stop) 
        $ResultNextLink = $Result."@odata.nextLink"
        $ServicePrincipals += $Result.value
    }

    return $ServicePrincipals
}

function Get-SignInEvents{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [String] $AppID
    )

    $uri = "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=AppId eq '$AppID'"
    $method = "GET"

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

    $SignIns = $Result.value

    $ResultNextLink = $Result."@odata.nextLink"
    while($ResultNextLink -ne $null){
        $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -ErrorAction Stop) 
        $ResultNextLink = $Result."@odata.nextLink"
        $SignIns += $Result.value
    }

    return $SignIns
}

# Get an Access Token
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

# Get all applications
$Applications = Get-Applications -Token $vars.Token.AccessToken 
$ServicePrincipals = Get-ServicePrincipals -Token $vars.Token.AccessToken

# Initialise variables to store sign-in statistics
$AllSignInEventApplications = @()
$AllSignInEventServicePrincipals = @()

# Get Application Sign-Ins
foreach($app in $Applications){
    Write-Host "Getting Sign-in events for $($app.displayName)"
    $signInEvents = Get-SignInEvents -Token $vars.token.AccessToken -AppID $app.appId

    $details = [ordered]@{}
    $details.add("App DisplayName", $app.displayName)
    $details.add("App ID", $app.appId)
    $details.add("Total SignIn Events", $signInEvents.count)
    $AllSignInEventApplications += New-Object PSObject -Property $details
}

# Get Service Principal Sign-Ins
foreach($app in $ServicePrincipals){
    Write-Host "Getting Sign-in events for $($app.displayName)"
    $signInEvents = Get-SignInEvents -Token $vars.token.AccessToken -AppID $app.appId

    $details = [ordered]@{}
    $details.add("App DisplayName", $app.displayName)
    $details.add("App ID", $app.appId)
    $details.add("Total SignIn Events", $signInEvents.count)
    $AllSignInEventServicePrincipals += New-Object PSObject -Property $details
}

# Export Sign-Ins to CSV
$AllSignInEventApplications | export-csv -NoTypeInformation $vars.Export.AppCSVLoc
$AllSignInEventServicePrincipals | export-csv -NoTypeInformation $vars.Export.SPCSVLoc

