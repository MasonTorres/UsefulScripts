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
        CSVLoc = "Sign-InEvents-All.csv"
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

function Get-SignInEventsAll{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token
    )

    $oneSuccessfulFetch = $False
    $retryCount = 0
    
    $uri = "https://graph.microsoft.com/beta/auditLogs/signIns?&`$filter=signInEventTypes/any(t: t eq 'nonInteractiveUser' or t eq 'servicePrincipal' or t eq 'managedIdentity' or t eq 'interactiveUser')"
    $method = "GET"

    $ResultNextLink = $uri
    while($ResultNextLink -ne $null){
        Try {
            $Result = (Invoke-RestMethod -Uri $ResultNextLink –Headers @{Authorization = "Bearer $token"} –Method $method -ErrorAction Stop) 
            $ResultNextLink = $Result."@odata.nextLink"
            $SignIns += $Result.value
            $oneSuccessfulFetch = $True
        }
        Catch [System.Net.WebException] {
            $statusCode = [int]$_.Exception.Response.StatusCode
            Write-Output $statusCode
            Write-Output $_.Exception.Message
            if($statusCode -eq 401 -and $oneSuccessfulFetch)
            {
                # Token might have expired! Renew token and try again
                $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                
                $oneSuccessfulFetch = $False
            }
            elseif($statusCode -eq 429)
            {
                # throttled request, wait for a few seconds and retry
                [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0])
                Write-Verbose -Message "Retry Caught, delaying $delay s"
                Start-Sleep -s $delay
            }
            elseif($statusCode -eq 403 -or $statusCode -eq 400 -or $statusCode -eq 401)
            {
                Write-Output "Please check the permissions of the user"
                break;
            }
            else {
                if ($retryCount -lt 5) {
                    Write-Output "Retrying..."
                    $retryCount++
                }
                else {
                    Write-Output "Download request failed. Please try again in the future."
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
    
                if ($retryCount -lt 5) {
                Write-Output "Retrying..."
                $retryCount++
            }
            else {
                Write-Output "Download request failed. Please try again in the future."
                break
            }
        }
    }

    return $SignIns
}

# Get an Access Token
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

# Get all Sign-in events
$signInEvents = Get-SignInEventsAll -Token $vars.token.AccessToken

# Initialise variables to store sign-in statistics
$AllSignInEvents = @()

foreach($signinEvent in $signinEvents){
    $details = [ordered]@{}
    $details.add("createdDateTime", $signinEvent.createdDateTime)
    $details.add("App DisplayName", $signinEvent.appDisplayName)
    $details.add("App ID", $signinEvent.appId)
    $details.add("Service Principal Name", $signinEvent.servicePrincipalName)
    $details.add("UPN", $signinEvent.userPrincipalName)
    $details.add("signInEventTypes", $signinEvent.signInEventTypes -join ";")
    $details.add("status.errorCode", $signinEvent.status.errorCode)
    $details.add("status.failureReason", $signinEvent.status.failureReason)
    $details.add("status.additionalDetails", $signinEvent.status.additionalDetails)
    $AllSignInEvents += New-Object PSObject -Property $details
}

# Export Sign-Ins to CSV
$AllSignInEvents | export-csv -NoTypeInformation $vars.Export.CSVLoc