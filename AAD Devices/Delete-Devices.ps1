#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

# 
# Create an App Registration 
# Assign Application permissions: Directory.Read.All, Directory.ReadWrite.All, Directory.AccessAsUser.All https://docs.microsoft.com/en-us/graph/api/device-list?view=graph-rest-1.0&tabs=http
# Assign Delegate permissions: 	Directory.AccessAsUser.All https://docs.microsoft.com/en-us/graph/api/device-delete?view=graph-rest-1.0&tabs=http
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

# Function: Get all devices using Microsoft Graph API
function Get-Devices{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token
    )

    $uri = "https://graph.microsoft.com/beta/devices"
    $method = "GET"

    $Devices = Invoke-MSGraph -Token $Token -Uri $uri -Method $method 

    return $Devices
}

# Function: Delete a signle device using Microsoft Graph API
function Invoke-DeleteDevice{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [string] $DeviceID
    )

    $uri = "https://graph.microsoft.com/beta/devices/$DeviceID"
    $method = "DELETE"

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method

    return $Result
}

# Function: Delete a signle device using Microsoft Graph API
function Invoke-DeleteDevicesBatch{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [hashtable] $batchRequest
    )

    $uri = "https://graph.microsoft.com/v1.0/`$batch"
    $method = "POST"

    $body = $batchRequest | convertTo-Json

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -Body $body

    return $Result
}

# Get our Access Token for Microsoft Graph API
$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret

# Get all risky users from Microsoft Graph API
$AllDevices = Get-Devices -Token $vars.Token

# Create a Hashtable with unique device names, then add any duplicate devices as child objects.
$devicesToCheck = @{}
foreach($device in $AllDevices){
    if($devicesToCheck[$device.displayName]){
        $devicesToCheck[$device.displayName] +=  @{
            $device.id = @{
                deviceId = $device.deviceId
                objectId = $device.id
                registrationDateTime = $device.registrationDateTime
                createdDateTime = $device.createdDateTime
                displayName = $device.displayName
                deleteStatus = ""
            }
        }
    }else{
        $devicesToCheck.add($device.displayName, @{
            $device.id = @{
                deviceId = $device.deviceId
                objectId = $device.id
                registrationDateTime = $device.registrationDateTime
                createdDateTime = $device.createdDateTime
                displayName = $device.displayName
                deleteStatus = ""
            }
        })
    }
}

# Loop through each unique device name, check the child devices and mark the most recently 
# created device as "Do Not Delete"
foreach($device in $devicesToCheck.GetEnumerator()){

    $latestActivity = get-date "1970-1-1"
        $latestActivityID = $null
    foreach($selectedDevice in $device.value.GetEnumerator()){

        foreach($duplicateDevice in $selectedDevice.value){
            if($duplicateDevice.createdDateTime -gt $latestActivity){
                $latestActivity = $duplicateDevice.createdDateTime
                $latestActivityID = $duplicateDevice.objectId
            }
        }

    }
    $devicesToCheck[$device.name][$latestActivityID].deleteStatus = "Do Not Delete"
}

# Export the hashtable as a JSON object
$jsonOutput =  $devicesToCheck | ConvertTo-Json
$jsonOutput | Out-File DuplicateDevices.json

# Export the hashtable as a CSV file
$CsvOutput = @()
foreach($device in $devicesToCheck.GetEnumerator()){
    
    foreach($deviceRegistration in $device.value){

        foreach($identifiedDevice in $deviceRegistration.GetEnumerator()){
            $details = [ordered]@{}
            $details.add("Computer", $device.Name)
            $details.add("DeleteStatus", $identifiedDevice.value.deleteStatus)
            $details.add("CreatedDateTime", $identifiedDevice.value.createdDateTime)
            $details.add("ObjectID", $identifiedDevice.value.objectId)
            $details.add("DeviceID", $identifiedDevice.value.deviceId)

            $CsvOutput+= New-Object PSObject -Property $details
        }
    }
}

$CsvOutput | Export-CSV DuplicateDevices.csv -NoTypeInformation

# Create batch requests and delete devices
foreach($device in $devicesToCheck.GetEnumerator()){
    
    # Create an array of batches
    # Each batch contains up to 20 Microsoft Graph API requests https://docs.microsoft.com/en-us/graph/json-batching
    [System.Collections.ArrayList]$batchRequestItemsArray = @()
    foreach($deviceRegistration in $device.value){
        $batchRequestItems = @()
        $count = 1
        foreach($identifiedDevice in $deviceRegistration.GetEnumerator()){
            if($identifiedDevice.Value.deleteStatus -ne "Do Not Delete" -and $count -le 20){
                $details = [ordered]@{}
                $details.add("id", $count)
                $details.add("method", "DELETE")
                $details.add("url", "/devices/$($identifiedDevice.Value.objectId)")

                $batchRequestItems += New-Object PSObject -Property $details
                $count++
            }
            
            if($count -eq 21 -and $identifiedDevice.Value.deleteStatus -ne "Do Not Delete"){
                $null = $batchRequestItemsArray.add($batchRequestItems)
                #$batchRequest.requests += $batchRequestItems
                $count = 1
                $batchRequestItems = @()
            }elseif( $batchRequestItemsArray.count * 20 + $count -eq $deviceRegistration.count){
                $null = $batchRequestItemsArray.add($batchRequestItems)
                $count = 1
            }
        }
    }

    # Loop through all the batches and execute batch request
    foreach($batch in $batchRequestItemsArray){
        if($batch.count -gt 0){
            #refresh token after 45mins time has elapsed
            if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
                Write-Output "Refreshing tokens" 

                $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                $vars.ScriptStartTime = Get-Date
            }

            # Execute batch job and delete devices
            $batchRequest = @{"requests" = $batch}
            Invoke-DeleteDevicesBatch -Token $vars.DelegateToken -batchRequest $batchRequest
        }
    }
}
