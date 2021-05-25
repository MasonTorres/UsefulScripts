#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

# Create an App Registration 
# Assign Application permissions: Device.ReadWrite.All https://docs.microsoft.com/en-us/graph/api/device-post-devices?view=graph-rest-1.0&tabs=http
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
    DelegateToken = @{

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

function Get-DelegateToken{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $clientId,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $clientSecret,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $tenantId,
        [Parameter(Mandatory=$false, Position=0)]
        [string] $refreshToken
    )
    
    # Original source from https://joymalya.com/powershell-ms-graph-api-part-1/
    #Define Client Variables Here
    #############################
    $resource = "https://graph.microsoft.com"
    $scope = "https://graph.microsoft.com/Directory.AccessAsUser.All"
    $redirectUri = "https://localhost"
    
    #UrlEncode variables for special characters
    ###########################################
    Add-Type -AssemblyName System.Web
    $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
    $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope)
    
    if( $null -eq $refreshToken -or $refreshToken.length -eq 0){
        #Obtain Authorization Code
        ##########################
        Add-Type -AssemblyName System.Windows.Forms
        $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
        $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }
        $url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded"
        $DocComp  = {
                $Global:uri = $web.Url.AbsoluteUri        
                if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
            }
        $web.ScriptErrorsSuppressed = $true
        $web.Add_DocumentCompleted($DocComp)
        $form.Controls.Add($web)
        $form.Add_Shown({$form.Activate()})
        $form.ShowDialog() | Out-Null
        $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
        $output = @{}
        foreach($key in $queryOutput.Keys){
            $output["$key"] = $queryOutput[$key]
        }
        $regex = '(?<=code=)(.*)(?=&)'
        $authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

        #Get Access Token with obtained Auth Code
        #########################################
        $body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource&scope=$scopeEncoded"
        $authUri = "https://login.microsoftonline.com/common/oauth2/token"
        $tokenResponse = Invoke-RestMethod -Uri $authUri -Method Post -Body $body -ErrorAction STOP

        $refreshToken = $null
    }else{

        #Get Access Token with obtained Auth Code
        #########################################
        $body = "grant_type=refresh_token&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&refresh_token=$refreshToken&resource=$resource&scope=$scopeEncoded"
        $authUri = "https://login.microsoftonline.com/common/oauth2/token"
        $tokenResponse = Invoke-RestMethod -Uri $authUri -Method Post -Body $body -ErrorAction STOP

        $refreshToken = $null
    }
    
    
    return @{
        "AccessToken" = $tokenResponse.access_token
        "RefreshToken" = $tokenResponse.refresh_token
        "IdToken" = $tokenResponse.id_token
    }
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
            $x = $_
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

function Invoke-CreateDevice{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [string] $DeviceID,
        [string] $key,
        [string] $date,
        [string] $name
    )

    $uri = "https://graph.microsoft.com/beta/devices"
    $method = "POST"

    $body = @{
            "accountEnabled" = $true
            "alternativeSecurityIds" = @(
                @{
                    "type" = 2
                    "key" = $key
                }
            )
            "deviceId" = $DeviceID
            "displayName" = $name
            "operatingSystem" = "Windows"
            "operatingSystemVersion" = "10.0.19042.985"
            "approximateLastSignInDateTime"= $date
    } | convertTo-Json

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -body $body

    return $Result
}

Function Get-RandomAlphanumericString {
	
	[CmdletBinding()]
	Param (
        [int] $length = 8
	)

	Begin{
	}

	Process{
        Write-Output ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) | Get-Random -Count $length  | % {[char]$_}) )
	}	
}

# Get our Access Token for Microsoft Graph API
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret


for($y = 0; $y -lt 50; $y++){
    for($x =1; $x -lt 15; $x++){

        #refresh token after 45mins time has elapsed
        if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
            Write-Output "Refreshing tokens" 

            $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
            $vars.ScriptStartTime = Get-Date
        }

        $randString = Get-RandomAlphanumericString -length 512
        $hasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256')
        $hash = $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($randString))
        $key = [System.Convert]::ToBase64String($hash)
        $name = "DuplicateDevice$y"

        $date = $(get-date $(Get-Date).addDays(-$x) -UFormat '+%Y-%m-%dT%H:%M:%S.000Z')
        Write-Host "Creating device $name" -ForegroundColor magenta
        Invoke-CreateDevice -Token $vars.DelegateToken -DeviceID (New-GUID).Guid -key $key -date $date -name $name

        Start-Sleep -s 2
    }
}
