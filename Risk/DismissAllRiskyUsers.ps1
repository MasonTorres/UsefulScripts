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

# Function: Manage calls to Microsoft Graph
# handles throttling and expired Access Tokens
function Invoke-MSGraph{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [Parameter(Mandatory=$true)]
         [string] $Uri,
         [Parameter(Mandatory=$true)]
         [string] $Method,
         [Parameter(Mandatory=$false)]
         [string] $Body
    )

    $ReturnValue = $null
    $OneSuccessfulFetch = $null
    $PermissionCheck = $false
    $RetryCount = 0
    $ResultNextLink = $Uri

    while($null -ne $ResultNextLink){
        Try {
            if($Body){
                $Result = Invoke-RestMethod -Method $Method -uri $ResultNextLink -ContentType "application/json" -Headers @{Authorization = "Bearer $($token.AccessToken)"} -Body $Body -ErrorAction Stop
            }else{
                $Result = Invoke-RestMethod -Method $Method -uri $ResultNextLink -ContentType "application/json" -Headers @{Authorization = "Bearer $($token.AccessToken)"} -ErrorAction Stop 
            }
            $ResultNextLink = $Result."@odata.nextLink"
            $ReturnValue += $Result.value
            $OneSuccessfulFetch = $true
        } 
        Catch [System.Net.WebException] {
            $statusCode = [int]$_.Exception.Response.StatusCode
            Write-Output "HTTP Status Code: $statusCode"
            Write-Output $_.Exception.Message
            if($statusCode -eq 401 -and $OneSuccessfulFetch)
            {
                Write-Output "HTTP Status Code: 401 - Trying to get new Access Token"
                # Token might have expired! Renew token and try again
                $Token.AccessToken = Get-AppToken -tenantId $Token.TenantID -clientId $Token.ClientID -clientSecret $Token.ClientSecret

                $OneSuccessfulFetch = $false
            }
            elseif ($statusCode -eq 401 -and $PermissionCheck -eq $false) {
                # In the case we are making multiple individual calls to Invoke-MSGraph we may need to check the access token has expired in between calls.
                # i.e the check above never occurs if MS Graph returns only one page of results.
                Write-Output "Retrying..."
                $Token.AccessToken = Get-AppToken -tenantId $Token.TenantID -clientId $Token.ClientID -clientSecret $Token.ClientSecret
                $PermissionCheck = $true
            }
            elseif($statusCode -eq 429)
            {
                Write-Output "HTTP Status Code: 429 - Throttled! Waiting some seconds"

                # throttled request, wait for a few seconds and retry
                [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0])
                Write-Verbose -Message "Retry Caught, delaying $delay s"
                Start-Sleep -s $delay
            }
            elseif($statusCode -eq 403 -or $statusCode -eq 400 -or $statusCode -eq 401)
            {
                Write-Output "Please check the permissions"
                break;
            }
            else {
                if ($RetryCount -lt 5) {
                    Write-Output "Retrying..."
                    $RetryCount++
                }
                else {
                    Write-Output "Request failed. Please try again in the future."
                    break
                }
            }
        }
        Catch {
            $exType = $_.Exception.GetType().FullName
            $exMsg = $_.Exception.Message
    
            Write-Output "Exception: $_.Exception"
            Write-Output "Error Message: $exType"
            Write-Output "Error Message: $exMsg"
    
            if ($RetryCount -lt 5) {
                Write-Output "Retrying..."
                $RetryCount++
            }
            else {
                Write-Output "Request failed. Please try again in the future."
                break
            }
        }
    }

    return $ReturnValue
}

# Function: Get all risky users using Microsoft Graph API
function Get-RiskyUsers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token
    )

    $uri = "https://graph.microsoft.com/beta/identityProtection/riskyUsers"
    $method = "GET"

    $RiskyUsers = Invoke-MSGraph -Token $Token -Uri $uri -Method $method 

    return $RiskyUsers
}

# Function: Dismiss risky users using Microsoft Graph API
function Invoke-DismissRiskyUsers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [system.array] $UserIds
    )

    $uri = "https://graph.microsoft.com/beta/identityProtection/riskyUsers/dismiss"
    $method = "POST"

    $body = @{
        "userIds" = @(
          $userIds
        )
      } | convertTo-Json

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -body $body

    return $Result
}

# Get our Access Token for Microsoft Graph API
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

# Get all risky users from Microsoft Graph API
$AllRiskyUsers = Get-RiskyUsers -Token $vars.Token

# Create a list of risky users whose riskState is not 'dismissed'
$riskyUsersToDissmis = @()
foreach($riskyUser in $AllRiskyUsers){
    if($riskyUser.riskState -ne 'dismissed'){
        $riskyUsersToDissmis+= $riskyUser.id
    }
}

# Dismiss risky users
Invoke-DismissRiskyUsers -Token $vars.Token -UserIds $riskyUsersToDissmis