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
# Assign Microsoft Graph Application Permission: Groups.ReadAll
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
    UserIDToCheck = ""
    ScriptStartTime = Get-Date
}

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

function Get-Groups{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [string] $UserID,
         [string] $SecurityEnabledOnly
    )

    $uri = "https://graph.microsoft.com/beta/users/$UserID/getMemberObjects"
    $method = "POST"

    $body = "
    {
      ""securityEnabledOnly"": ""$SecurityEnabledOnly""
    }
    "

    $Result = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -body $body -ErrorAction Stop
    
    $Groups = $Result.value

    $ResultNextLink = $Result."@odata.nextLink"
    while($ResultNextLink -ne $null){
        $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -ErrorAction Stop) 
        $ResultNextLink = $Result."@odata.nextLink"
        $Groups += $Result.value
    }

    return $Groups
}

$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

$group1 = Get-Groups -Token $vars.Token.AccessToken -SecurityEnabledOnly "true" -UserID $vars.UserIDToCheck
$group2 = Get-Groups -Token $vars.Token.AccessToken -SecurityEnabledOnly "false" -UserID $vars.UserIDToCheck

$GroupsNotIncludedInClaims = @()

foreach($group2Members in $group2){
    $isInGroup=$false
    foreach($group1Members in $group1){
        if($group2Members -eq $group1Members){
            $isInGroup=$true
        }
    }

    if($isInGroup -eq $false){
        $GroupsNotIncludedInClaims+=$group2Members
    }
}

Write-Host "$($GroupsNotIncludedInClaims.count) groups omitted from claim"
foreach($omittedGroup in $GroupsNotIncludedInClaims){
    Write-Host "Group: $omittedGroup"
}