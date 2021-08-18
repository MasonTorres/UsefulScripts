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
# Assign Delegate permissions: 	Directory.AccessAsUser.All https://docs.microsoft.com/en-us/graph/api/device-delete?view=graph-rest-1.0&tabs=http
# Create Application secret
# Set the App Registration to Public
# Update $vars variable below: ClientSecret, ClientID and TenantID 

# This script will delete devices from Azure AD.
# App ClientSecret, ClientID and TenantID are needed for delegate token refresh

# Kudos to https://blog.simonw.se/getting-an-access-token-for-azuread-using-powershell-and-device-login-flow/
# Kudos to https://github.com/Azure-Samples/DSRegTool

param (
    [Parameter( ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true,
                HelpMessage="Clean up stale Azure AD devices")]
    [string]$ClientID = "",
    [string]$ClientSecret = "",
    [string]$TenantID = ""
)

# Variable: Store variables used throughout this script
$vars = @{
    # Used to generate an Access Token to query Microsoft Graph API.
    Token = @{
        ClientSecret = $ClientSecret
        ClientID = $ClientID
        TenantID = $TenantID
        AccessToken = ""
    }
    DelegateToken = @{}
    ScriptStartTime = Get-Date

    AllDevices = @{}
    DevicesToCheck = @{}

    AllUsers = @{
        FromGraph = @{}
        WithADevice = @{}
        WithNoDevice = @{}
    }
    

    LogFile = "CleanUp-StaleDevicesLog.txt"
}

Add-Type -AssemblyName System.Web

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

# Function: Create Access Token with delegated permissions using device code flow
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
    
    # Original source from https://blog.simonw.se/getting-an-access-token-for-azuread-using-powershell-and-device-login-flow/
    $resource = "https://graph.microsoft.com"
    $Timeout = 300
    $scope =  "https://graph.microsoft.com/.Default"
    $redirectUri = "https://localhost"

    $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
    $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
    $resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
    $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope)


    if( $null -eq $refreshToken -or $refreshToken.length -eq 0){

        $DeviceCodeRequestParams = @{
            Method = 'POST'
            Uri    = "https://login.microsoftonline.com/$TenantID/oauth2/devicecode"
            Body   = @{
                client_id = $ClientId
                resource  = $Resource
            }
        }

        $DeviceCodeRequest = Invoke-RestMethod @DeviceCodeRequestParams
        Write-Host $DeviceCodeRequest.message -ForegroundColor Yellow

        $TokenRequestParams = @{
            Method = 'POST'
            Uri    = "https://login.microsoftonline.com/$TenantId/oauth2/token"
            Body   = @{
                grant_type = "urn:ietf:params:oauth:grant-type:device_code"
                code       = $DeviceCodeRequest.device_code
                client_id  = $ClientId
            }
        }

        $TimeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()
        while ([string]::IsNullOrEmpty($TokenRequest.access_token)) {
            if ($TimeoutTimer.Elapsed.TotalSeconds -gt $Timeout) {
                throw 'Login timed out, please try again.'
            }
            $TokenRequest = try {
                Invoke-RestMethod @TokenRequestParams -ErrorAction Stop
            }
            catch {
                $Message = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($Message.error -ne "authorization_pending") {
                    throw
                }
            }
            Start-Sleep -Seconds 1
        }

        $tokenResponse = $TokenRequest

    }else{

        #Use RefreshToken to acquire a new Access Token
        try{
            $body = "grant_type=refresh_token&redirect_uri=$redirectUriEncoded&client_id=$clientId&client_secret=$clientSecretEncoded&refresh_token=$refreshToken&resource=$resourceEncoded&scope=$scopeEncoded"
            $authUri = "https://login.microsoftonline.com/common/oauth2/token"
            $tokenResponse = Invoke-RestMethod -Uri $authUri -Method Post -Body $body -ErrorAction STOP
        }catch{
            $e = $_
            Write-Output "Error refreshing token. Acquire a new token"
            throw
        }
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
         [string] $Body,
         [Parameter(Mandatory=$false)]
         [bool] $ShowProgress = $true,
         [Parameter(Mandatory=$false)]
         [bool] $useConsistencyLevel = $false,
         [Parameter(Mandatory=$false)]
         [bool] $DebugLots = $false
    )

    Write-Log "Entering Graph call Uri: $Uri Method: $Method"
    if($DebugLots){
        Write-Log "Body: $body"
    }

    $ReturnValue = $null
    $OneSuccessfulFetch = $null
    $PermissionCheck = $false
    $RetryCount = 0
    $ResultNextLink = $Uri
    $batchCount = 0

    $headers = @{
        Authorization = "Bearer $($token.AccessToken)"
    }

    if($useConsistencyLevel -eq $true){
        $headers.add("ConsistencyLevel", "eventual");
    }

    while($null -ne $ResultNextLink){
        Try {
            if($Body){
                $Result = Invoke-RestMethod -Method $Method -uri $ResultNextLink -ContentType "application/json" -Headers $headers -Body $Body -ErrorAction Stop
            }else{
                $Result = Invoke-RestMethod -Method $Method -uri $ResultNextLink -ContentType "application/json" -Headers $headers -ErrorAction Stop 
            }
            $ResultNextLink = $Result."@odata.nextLink"
            $ReturnValue += $Result.value
            $OneSuccessfulFetch = $true
            $batchCount++
            if($ShowProgress -eq $true){
                Write-Host -NoNewline "`rRecevied pages from Microsoft Graph API $batchCount"
            }
            if($result.responses){
                $ReturnValue = $result
            }
        } 
        Catch [System.Net.WebException] {
            $x = $_
            $statusCode = [int]$_.Exception.Response.StatusCode
            Write-Log "HTTP Status Code: $statusCode"
            Write-Log $_.Exception.Message
            Write-Output $_.Exception.Message
            if($statusCode -eq 401 -and $OneSuccessfulFetch)
            {
                Write-Log "HTTP Status Code: 401 - Trying to get new Access Token"
                # Token might have expired! Renew token and try again
                try{
                    Write-Log "Refreshing tokens within Graph call." 
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                    $vars.ScriptStartTime = Get-Date
                    Write-Log "Success"
                    
                }catch{
                    Write-Log "Could not refresh token."
                    Write-Log $_.Exception.Message
                    Write-Log "Trying to get new token"
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }

                $OneSuccessfulFetch = $false
            }
            elseif ($statusCode -eq 401 -and $PermissionCheck -eq $false) {
                # In the case we are making multiple individual calls to Invoke-MSGraph we may need to check the access token has expired in between calls.
                # i.e the check above never occurs if MS Graph returns only one page of results.
                Write-Log "Retrying...Getting new access token"
                try{
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                    $vars.ScriptStartTime = Get-Date
                    
                }catch{
                    $e = $_
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }
                
                $PermissionCheck = $true
            }
            elseif($statusCode -eq 429)
            {
                Write-Log "HTTP Status Code: 429 - Throttled! Waiting some seconds"

                # throttled request, wait for a few seconds and retry
                try{
                    [int] $delay = [int](($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0])
                    Write-Log "Retry Caught, delaying $delay s"
                }catch{
                    #Error receiving delay in seconds. Set to 5 seconds.
                    $delay = 5
                }
                Start-Sleep -s $delay
            }
            elseif($statusCode -eq 403 -or $statusCode -eq 400 -or $statusCode -eq 401)
            {
                Write-Log "Please check the permissions"
                break;
            }
            else {
                if ($RetryCount -lt 5) {
                    Write-Log "Retrying...Retry count: $RetryCount"
                    $RetryCount++
                }
                else {
                    Write-Log "Request failed. Please try again in the future."
                    break
                }
            }
        }
        Catch {
            $exType = $_.Exception.GetType().FullName
            $exMsg = $_.Exception.Message
    
            Write-Log "Exception: $($_.Exception)"
            Write-Log "Error Message: $exType"
            Write-Log "Error Message: $exMsg"
    
            if ($RetryCount -lt 5) {
                Write-Log "Retrying...Retry count: $RetryCount"
                $RetryCount++
            }
            else {
                Write-Log "Graph request failed. Please try again in the future."
                break
            }
        }

        #hot fix :D
        #refresh token after 45mins time has elapsed
        if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
            Write-Log "45 mins elapsed. StartTime: Get-Date $($vars.ScriptStartTime)"

            try{
                Write-Log "Refreshing tokens within Graph call." 
                $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                $vars.ScriptStartTime = Get-Date
                Write-Log "Success" 
                
            }catch{
                Write-Log "Could not refresh token."
                Write-Log $_.Exception.Message
                Write-Log "Trying to get new token"
                $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
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

    $uri = "https://graph.microsoft.com/beta/devices?`$top=999&select=registrationDateTime,createdDateTime,displayName,id,deviceid,deviceOwnership,approximateLastSignInDateTime"
    $method = "GET"

    $Devices = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -DebugLots $true

    return $Devices
}

function Get-Users{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token
    )

    $uri = "https://graph.microsoft.com/beta/users?`$top=999&select=userprincipalname,id"
    $method = "GET"

    $Users = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -DebugLots $true

    return $Users
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

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -DebugLots $false 

    return $Result
}

# Function: Delete a signle device using Microsoft Graph API
function Invoke-BatchRequest{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [hashtable] $batchRequest,
         [Parameter(Mandatory=$false)]
         [bool] $useConsistencyLevel = $false
    )

    $uri = "https://graph.microsoft.com/beta/`$batch"
    $method = "POST"

    $body = $batchRequest | convertTo-Json

    $Result = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -Body $body -ShowProgress $false -useConsistencyLevel $useConsistencyLevel -DebugLots $false

    return $Result
}

function Invoke-CheckVariables{
    while([string]::IsNullOrEmpty($vars.token.ClientSecret) -or [string]::IsNullOrEmpty($vars.token.ClientID) -or [string]::IsNullOrEmpty($vars.token.TenantID)){
        Write-Host "Client ID, Secret or Tenant ID is missing. Please enter:"
        while([string]::IsNullOrEmpty($vars.token.ClientID)){
            $vars.token.ClientID = Read-Host -Prompt "Enter Client ID"
        }

        while([string]::IsNullOrEmpty($vars.token.ClientSecret)){
            $vars.token.ClientSecret = Read-Host -Prompt "Enter Client Secret"
        }

        while([string]::IsNullOrEmpty($vars.token.TenantID)){
            $vars.token.TenantID = Read-Host -Prompt "Enter Tenant ID"
        }
    }

    try{
        Write-Log "Getting new Acess Token using Refresh Token"
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
        $vars.ScriptStartTime = Get-Date
        Write-Log "Success"
        
    }catch{
        $e = $_
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    }

    Clear-Host
}

Function Write-Log{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $String
    )
    $datetimeUTC = get-date -f u
    Add-Content $vars.LogFile "$datetimeUTC $string"
}

Function Step1{
    # Get our Access Token for Microsoft Graph API
    try{
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
        $vars.ScriptStartTime = Get-Date
        
    }catch{
        $e = $_
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    }

    ''
    Write-Log "Getting devices..."
    Write-Host "Getting devices..."

    # Get all risky users from Microsoft Graph API
    $vars.AllDevices = Get-Devices -Token $vars.DelegateToken

    $deviceCheckCount = 0
    Write-Log "Evaluating devices. This may take some time."
    Write-Host "`nEvaluating devices. This may take some time. Please do not close the terminal session."
    # Create a Hashtable with unique device names, then add any duplicate devices as child objects.
    $vars.DevicesToCheck = @{}
    foreach($device in $vars.AllDevices){
        $deviceCheckCount++
        $percent = [math]::Round($(($deviceCheckCount / $($vars.AllDevices.count)) * 100),2)
        Write-Host -NoNewline "`rEvaluating devices $deviceCheckCount / $($vars.AllDevices.count) $percent% "

        if($vars.DevicesToCheck[$device.displayName]){
            
            $vars.DevicesToCheck[$device.displayName] +=  @{
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
            $vars.DevicesToCheck.add($device.displayName, @{
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
    foreach($device in $vars.DevicesToCheck.GetEnumerator()){

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
        $vars.DevicesToCheck[$device.name][$latestActivityID].deleteStatus = "Do Not Delete"
    }

    Write-Log "Evaluating devices. Complete."

    ''
    Write-Log "Number of Devices: $($vars.AllDevices.count)"
    Write-Host "Number of Devices: $($vars.AllDevices.count)" -ForegroundColor Magenta
    Write-Log "Number of Unique devices by name: $($vars.DevicesToCheck.count)"
    Write-Host "Number of Unique devices by name: $($vars.DevicesToCheck.count)" -ForegroundColor Magenta
    ''
}

Function Step1A{
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        Write-Log "No devices found. please run option 1, 2 or 3"
        Write-Host "No devices found. please run option 1, 2 or 3" -ForegroundColor Yellow
    }else{
        try{
            # Export the hashtable as a CSV file
            # $CsvOutput = @()
            $CsvOutput = [System.Text.StringBuilder]::new()
            [void]$CsvOutput.AppendLine( "Computer,DeleteStatus,CreatedDateTime,ObjectID,DeviceID,userPrincipalName" )
            $count = 0;
            $countTotal = $vars.DevicesToCheck.count
               
            foreach($device in $vars.DevicesToCheck.GetEnumerator()){
                Write-Host -NoNewline "`rCreating CSV export devices  $count / $countTotal"
                $count++

                foreach($deviceRegistration in $device.value){

                    foreach($identifiedDevice in $deviceRegistration.GetEnumerator()){
                        # $details = [ordered]@{}
                        # $details.add("Computer", $device.Name)
                        # $details.add("DeleteStatus", $identifiedDevice.value.deleteStatus)
                        # $details.add("CreatedDateTime", $identifiedDevice.value.createdDateTime)
                        # $details.add("ObjectID", $identifiedDevice.value.objectId)
                        # $details.add("DeviceID", $identifiedDevice.value.deviceId)

                        # $CsvOutput+= New-Object PSObject -Property $details

                        [void]$CsvOutput.AppendLine( "$($device.Name),$($identifiedDevice.value.deleteStatus),$($identifiedDevice.value.createdDateTime),$($identifiedDevice.value.objectId),$($identifiedDevice.value.deviceId),$($identifiedDevice.value.userPrincipalName)" )
                    }
                }
            }
            # $CsvOutput | Export-CSV DuplicateDevices.csv -NoTypeInformation
            $CsvOutput.ToString() | Out-File DuplicateDevices.csv -Encoding ascii
            Clear-Host
            Write-Log "Sucecssfully exported CSV file."
            Write-Host "Sucecssfully exported CSV file." -ForegroundColor Magenta
        }catch{
            Write-Log "An Error occurred exporting CSV file"
            Write-Host "An Error occurred exporting CSV file" -ForegroundColor Red
            Write-Error $_
        }
    }
}

Function Step1B{
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        Write-Log "No devices found. please run option 1, 2 or 3"
        Write-Host "No devices found. please run option 1, 2 or 3" -ForegroundColor Yellow
    }else{
        try{
            # Export the hashtable as a JSON object
            $jsonOutput =  $vars.DevicesToCheck | ConvertTo-Json
            $jsonOutput | Out-File DuplicateDevices.json
            Clear-Host
            Write-Log "Sucecssfully exported JSON file."
            Write-Host "Sucecssfully exported JSON file." -ForegroundColor Magenta
        }catch{
            Write-Log "An Error occurred exporting JSON file"
            Write-Host "An Error occurred exporting JSON file" -ForegroundColor Red
            Write-Error $_
        }
    }
}

Function Step2{
    Write-Log "Prompt user to select CSV file from Windows Explorer"
    Write-Host "Select CSV file from Windows Explorer"

    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToCsv = $OpenFileDialog.filename

    Clear-Host
    try{
        Write-Log "Importing CSV file"
        Write-Host "Importing CSV file"
        $csvImport = Import-CSV $pathToCsv

        $row = 1
        $vars.DevicesToCheck = @{}
        foreach($device in $csvImport){

            $row++
            if([string]::IsNullOrEmpty($device.Computer) -or [string]::IsNullOrEmpty($device.CreatedDateTime) -or [string]::IsNullOrEmpty($device.ObjectID) -or [string]::IsNullOrEmpty($device.DeviceID)){
                Write-Log "Error on row $row. Skipping row."
                Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
                if([string]::IsNullOrEmpty($device.Computer)){
                    Write-Log "Computer is empty"
                    Write-Host "Computer is empty"
                }

                if([string]::IsNullOrEmpty($device.CreatedDateTime)){
                    Write-Log "CreatedDateTime is empty"
                    Write-Host "CreatedDateTime is empty"
                }

                if([string]::IsNullOrEmpty($device.ObjectID)){
                    Write-Log "ObjectID is empty"
                    Write-Host "ObjectID is empty"
                }

                if([string]::IsNullOrEmpty($device.DeviceID)){
                    Write-Log "DeviceID is empty"
                    Write-Host "DeviceID is empty"
                }
                ''
            }else{
                if($vars.DevicesToCheck[$device.Computer]){
                
                    $vars.DevicesToCheck[$device.Computer] +=  @{
                        $device.ObjectID = @{
                            deviceId = $device.deviceId
                            objectId = $device.ObjectID
                            registrationDateTime = $device.registrationDateTime
                            createdDateTime = $device.createdDateTime
                            displayName = $device.Computer
                            deleteStatus = $device.DeleteStatus
                            userPrincipalName = $device.userPrincipalName
                        }
                    }
                }else{
                    $vars.DevicesToCheck.add($device.Computer, @{
                        $device.ObjectID = @{
                            deviceId = $device.deviceId
                            objectId = $device.ObjectID
                            registrationDateTime = $device.registrationDateTime
                            createdDateTime = $device.createdDateTime
                            displayName = $device.Computer
                            deleteStatus = $device.DeleteStatus
                            userPrincipalName = $device.userPrincipalName
                        }
                    })
                }
            }
        }

        Clear-Host
        Write-Log "Sucecssfully imported CSV file."
        Write-Host "Sucecssfully imported CSV file." -ForegroundColor Magenta
    }catch{
        WRite-Log "An Error occurred reading CSV file"
        Write-Host "An Error occurred reading CSV file" -ForegroundColor Red
        Write-Error $_
    }
}

Function Step3{
    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    Write-Log "Prompt user to select JSON file from Windows Explorer"
    Write-Host "Select JSON file from Windows Explorer"

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "JSON (*.json)| *.json"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToJson = $OpenFileDialog.filename

    Clear-Host
    Write-Log "Importing JSON file"
    Write-Host "Importing JSON file"

    try{
        $errorImportingDevice = 0
        $jsonImport = Get-Content $pathToJson | ConvertFrom-Json 

        $vars.DevicesToCheck = @{}
        foreach($deviceFromJson in $jsonImport.PSObject.Properties){

            $device = @{
                $deviceFromJson.Name = @{}
            }

            foreach($dupDevice in $deviceFromJson.value.psobject.Properties.value){

                if( [string]::IsNullOrEmpty($dupDevice.displayName) -or 
                    [string]::IsNullOrEmpty($dupDevice.createdDateTime) -or 
                    [string]::IsNullOrEmpty($dupDevice.objectId) -or 
                    [string]::IsNullOrEmpty($dupDevice.deviceId))
                {
                    Write-Host "Error on json value." -ForegroundColor Red
                    if([string]::IsNullOrEmpty($dupDevice.displayName)){
                        Write-Log "$($dupDevice.objectId): Computer is empty"
                        Write-Host "$($dupDevice.objectId): Computer is empty"
                        $errorImportingDevice++
                    }
        
                    if([string]::IsNullOrEmpty($dupDevice.createdDateTime)){
                        Write-Log "$($dupDevice.displayName): CreatedDateTime is empty"
                        Write-Host "$($dupDevice.displayName): CreatedDateTime is empty"
                        $errorImportingDevice++
                    }
        
                    if([string]::IsNullOrEmpty($dupDevice.objectId)){
                        Write-Log "$($dupDevice.displayName): ObjectID is empty"
                        Write-Host "$($dupDevice.displayName): ObjectID is empty"
                        $errorImportingDevice++
                    }
        
                    if([string]::IsNullOrEmpty($dupDevice.deviceId)){
                        Write-Log "$($dupDevice.displayName): DeviceID is empty"
                        Write-Host "$($dupDevice.displayName): DeviceID is empty"
                        $errorImportingDevice++
                    }
                    ''
                }else{

                    if($device[$deviceFromJson.Name]){
                        $device.($deviceFromJson.Name) += @{
                            $dupDevice.objectId = @{
                            displayName = $dupDevice.displayName
                            objectId = $dupDevice.objectId
                            createdDateTime = $dupDevice.createdDateTime
                            registrationDateTime = $dupDevice.registrationDateTime
                            deviceId = $dupDevice.deviceId
                            deleteStatus = $dupDevice.deleteStatus
                            userPrincipalName = $device.userPrincipalName
                        }
                    }
                    }else{
                        $device.add($deviceFromJson.Name, @{
                                $dupDevice.objectId = @{
                                displayName = $dupDevice.displayName
                                objectId = $dupDevice.objectId
                                createdDateTime = $dupDevice.createdDateTime
                                registrationDateTime = $dupDevice.registrationDateTime
                                deviceId = $dupDevice.deviceId
                                deleteStatus = $dupDevice.deleteStatus
                                userPrincipalName = $device.userPrincipalName
                            }
                        })
                    }
                }
            }

            $vars.DevicesToCheck += $device
        }

        if($errorImportingDevice -gt 0){
            Write-Log "There is missing data in some of the json values. Please fix source data and re-run"
            Write-Host "There is missing data in some of the json values. Please fix source data and re-run" -ForegroundColor Red
            $vars.DevicesToCheck = @{}
        }else{
            Clear-Host
            Write-Log "Sucecssfully imported JSON file."
            Write-Host "Sucecssfully imported JSON file." -ForegroundColor Magenta
        }
    }catch{
        Write-Log "An Error occurred reading JSON file" 
        Write-Host "An Error occurred reading JSON file" -ForegroundColor Red
        Write-Error $_
    }
}

Function Step4{
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        Write-Log "No devices found. please run option 1, 2 or 3" 
        Write-Host "No devices found. please run option 1, 2 or 3" -ForegroundColor Yellow
    }else{
        $numDevicesToDelete = 0
        $numDevicesNotToDelete = 0
        foreach($device in $vars.DevicesToCheck.GetEnumerator()){
            foreach($deviceDuplicate in $device.value){
                foreach($identifiedDevice in $deviceDuplicate.GetEnumerator()){
                    if($identifiedDevice.value.deleteStatus -ne "Do Not Delete"){
                        $numDevicesToDelete++
                    }else{
                        $numDevicesNotToDelete++
                    }
                }
            }
        }

        ''
        Write-Host "This action will delete $numDevicesToDelete devices and ignore $numDevicesNotToDelete devices" -ForegroundColor Red
        Write-Host "This action can NOT be undone!" -ForegroundColor Red
        ''
        $ProceedToDeleteDevices = Read-Host -Prompt "Are you sure you want to delete $numDevicesToDelete devices? (Y/N)"
        ''

        if($ProceedToDeleteDevices -eq "Y"){
            $numDevicesDeleted = 0

            # Create batch requests and delete devices
            foreach($device in $vars.DevicesToCheck.GetEnumerator()){
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
                            $numDevicesDeleted++
                        }
                        
                        #Add batch of 20 to array of batches
                        if($device.value.count -lt 20 -and $count -eq $device.value.count){
                            #There are less than 20 devices.  Adding batch to array of batches.
                            $null = $batchRequestItemsArray.add($batchRequestItems)
                        }elseif($count -eq 21 -and $identifiedDevice.Value.deleteStatus -ne "Do Not Delete"){
                            #We have reached 20 devices. Adding batch to array of batches.
                            $null = $batchRequestItemsArray.add($batchRequestItems)
                            $count = 1
                            $batchRequestItems = @()
                        }elseif( $batchRequestItemsArray.count * 20 + $count-1 -eq $deviceRegistration.count){
                            #Add the remaing batch. May be less than 20. Adding batch to array of batches.
                            $null = $batchRequestItemsArray.add($batchRequestItems)
                            $count = 1
                        }

                        Write-Host "`rDeleting devices $numDevicesDeleted / $numDevicesToDelete" -NoNewline
                    }
                }



                # Loop through all the batches and execute batch request
                foreach($batch in $batchRequestItemsArray){
                    if($batch.count -gt 0){
                        #refresh token after 45mins time has elapsed
                        if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
                            Write-Output "Refreshing tokens" 

                            try{
                                $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                                $vars.ScriptStartTime = Get-Date
                                
                            }catch{
                                $e = $_
                                $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                            }
                        }

                        # Execute batch job and delete devices
                        $batchRequest = @{"requests" = $batch}
                        $deleteResult = Invoke-BatchRequest -Token $vars.DelegateToken -batchRequest $batchRequest
                        
                    }
                }
            }
        }
    }
}

#User specific functions
Function Step5{
# Get our Access Token for Microsoft Graph API
    try{
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
        $vars.ScriptStartTime = Get-Date
        
    }catch{
        $e = $_
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    }

    ''
    Write-Log "Getting users..."
    Write-Host "Getting users..."

    # Get all users from Microsoft Graph API
    $vars.AllUsers.FromGraph = Get-Users -Token $vars.DelegateToken

    $AllUsersHash = @{}
    $vars.AllUsers.FromGraph | ForEach-Object{
        $AllUsersHash[$_.id] = $_.userprincipalname 
    }

    # Create the graph calls for each device and add to a list (batch).
    $userCount = 0
    $batchCount = 1
    [System.Collections.ArrayList]$batchRequestItemsArray = @()
    $batchRequestItems = @()
    $AllUsersCount = $vars.AllUsers.FromGraph.count
    Write-Host "`nCreating Batches..."
    Write-Log "Number of users: $AllUsersCount"
    Write-Log "Creating Batches..."
    foreach($user in $vars.AllUsers.FromGraph){
        $percent = [math]::Round($(($userCount / $AllUsersCount) * 100),2)
        $userCount++
        Write-Host -NoNewline "`rCreating batches of users $userCount / $AllUsersCount $percent %"

        $details = [ordered]@{}
        $details.add("id", $batchCount)
        $details.add("method", "GET")
        #$details.add("url", "/devices?`$search=""physicalIds:[USER-HWID]:$($user.id)""&`$select=registrationDateTime,createdDateTime,displayName,id,deviceid,approximateLastSignInDateTime,physicalIds&ConsistencyLevel=eventual")
        $details.add("url", "/users/$($user.id)/registeredDevices?`$select=registrationDateTime,createdDateTime,displayName,id,deviceid,approximateLastSignInDateTime,physicalIds,displayName")
        $batchRequestItems += New-Object PSObject $details
        $batchCount++

        if($userCount % 20 -eq 0){
            $null = $batchRequestItemsArray.add($batchRequestItems)
            $batchRequestItems = @()
            $batchCount = 1
        }elseif($userCount -eq $AllUsersCount){ # Make sure we get the last batch which may have less than 20 items
            $null = $batchRequestItemsArray.add($batchRequestItems)
            $batchRequestItems = @()
        }
    }

    $DeviceAndTheirUsers = @{}
    $UsersWithNoDevices = @{}
    $batchRunCount = 0
    $batchRequestItemsArrayCount = $batchRequestItemsArray.count
    Write-Host "`nExecuting batches..."
    Write-Log "Executing batches..."
    Write-Log "Number of batches: $batchRequestItemsArrayCount"
    # Loop through all the batches and execute batch request
    foreach($batch in $batchRequestItemsArray){
        $percent = [math]::Round($(($batchRunCount / $batchRequestItemsArrayCount) * 100),2)
        Write-Host -NoNewline "`rRunning batch  $batchRunCount / $batchRequestItemsArrayCount $percent %"
        $batchRunCount++

        if($batch.count -gt 0){
            #refresh token after 45mins time has elapsed
            if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
                Write-Output "Refreshing tokens" 

                try{
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                    $vars.ScriptStartTime = Get-Date
                    
                }catch{
                    $e = $_
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }
            }

            # Execute batch job
            $batchRequest = @{"requests" = $batch}
            $batchResult = Invoke-BatchRequest -Token $vars.DelegateToken -batchRequest $batchRequest -useConsistencyLevel $true  #$batchRequest
            
            # Loop through Graph results 
            foreach($individualResult in $batchResult.responses){

                if($individualResult.status -eq 200){
                    # Devices found for given user
                    if($individualResult.body.value.count -gt 0){
                        
                        foreach($device in $individualResult.body.value){
                            $userId = $($device.physicalIds[0]).split(":")[1]
                           
                            $DeviceAndTheirUsers[$device.id] = @{
                                deviceId = $device.id
                                registrationDateTime = $device.registrationDateTime
                                createdDateTime = $device.createdDateTime
                                approximateLastSignInDateTime = $device.approximateLastSignInDateTime
                                userprincipalname = $AllUsersHash[$userId]
                                displayName = $device.displayName
                            }
                        }
                    }else{
                        foreach($b in $batchRequest.requests){
                            if($b.id -eq $individualResult.id){
                                $userId = $($b.url).split('/')[2]
                            }
                        }
                        $UsersWithNoDevices[$userId] = @{
                            userId = $userId
                            userprincipalname = $AllUsersHash[$userId]
                        }
                    }                    
                }
            }
        }
    }
    
    $vars.AllUsers.WithNoDevice = $UsersWithNoDevices
    $vars.AllUsers.WithADevice = $DeviceAndTheirUsers

}

Function Step5A{
    if(([string]::IsNullOrEmpty($vars.AllUsers.WithNoDevice) -or ($vars.AllUsers.WithNoDevice.count -eq 0) ) -and ([string]::IsNullOrEmpty($vars.AllUsers.WithADevice) -or ($vars.AllUsers.WithADevice.count -eq 0))){
        Write-Log "No users found. Please run option 5"
        Write-Host "No users found. Please run option 5" -ForegroundColor Yellow
    }else{
        try{
            # Export the hashtable as a CSV file
            $CsvOutputUsersWithNoDevice = [System.Text.StringBuilder]::new()
            [void]$CsvOutputUsersWithNoDevice.AppendLine( "UserId,UserPrincipalName" )
            $countNoDevice = 0;
            $countNoDeviceTotal = $vars.AllUsers.WithNoDevice.count
            foreach($user in $vars.AllUsers.WithNoDevice.GetEnumerator()){
                Write-Host -NoNewline "`rCreating CSV export user with no device  $countNoDevice / $countNoDeviceTotal"
                $countNoDevice++

                [void]$CsvOutputUsersWithNoDevice.AppendLine("$($user.value.userId),$($user.value.userprincipalname)")
            }
            $CsvOutputUsersWithNoDevice.ToString() | Out-File DuplicateDevices-AllUsersWithNoDevices.csv -Encoding ascii

            $CsvOutputUsersWithADevice = [System.Text.StringBuilder]::new()
            [void]$CsvOutputUsersWithADevice.AppendLine( "UserPrincipalName,approximateLastSignInDateTime,createdDateTime,deviceId,registrationDateTime,displayName" )
            $countADevice = 0;
            $countADeviceTotal = $vars.AllUsers.WithADevice.count
            foreach($user in $vars.AllUsers.WithADevice.GetEnumerator()){
                Write-Host -NoNewline "`rCreating CSV export users with a device $countADevice / $countADeviceTotal"
                $countADevice++

                [void]$CsvOutputUsersWithADevice.AppendLine("$($user.value.UserPrincipalName),$($user.value.approximateLastSignInDateTime),$($user.value.createdDateTime),$($user.value.deviceId),$($user.value.registrationDateTime),$($user.value.displayName)")
            }
            $CsvOutputUsersWithADevice.ToString() | Out-File  DuplicateDevices-AllUsersWithADevices.csv -Encoding ascii

            Clear-Host
            Write-Log "Sucecssfully exported CSV file."
            Write-Host "Sucecssfully exported CSV file." -ForegroundColor Magenta
        }catch{
            Write-Log "An Error occurred exporting CSV file"
            Write-Host "An Error occurred exporting CSV file" -ForegroundColor Red
            Write-Error $_
        }
    }
}

Function Step5B{
    if(([string]::IsNullOrEmpty($vars.AllUsers.WithNoDevice) -or ($vars.AllUsers.WithNoDevice.count -eq 0) ) -and ([string]::IsNullOrEmpty($vars.AllUsers.WithADevice) -or ($vars.AllUsers.WithADevice.count -eq 0))){
        Write-Log "No devices found. please run option 5"
        Write-Host "No devices found. please run option 5" -ForegroundColor Yellow
    }else{
        try{
            # Export the hashtable as a JSON object
            $jsonOutputADevice =  $vars.AllUsers.WithADevice | ConvertTo-Json
            $jsonOutputADevice | Out-File DuplicateDevices-AllUsersWithADevices.json

            $jsonOutputNoDevice =  $vars.AllUsers.WithNoDevice | ConvertTo-Json
            $jsonOutputNoDevice | Out-File DuplicateDevices-AllUsersWithNoDevices.json
            Clear-Host
            Write-Log "Sucecssfully exported JSON file."
            Write-Host "Sucecssfully exported JSON file." -ForegroundColor Magenta
        }catch{
            Write-Log "An Error occurred exporting JSON file"
            Write-Host "An Error occurred exporting JSON file" -ForegroundColor Red
            Write-Error $_
        }
    }
}

Function Step6{
    #Users with no device
    Write-Log "Prompt user to select CSV file from Windows Explorer: All users with no device"
    Write-Host "Select CSV file from Windows Explorer: All users with no device"

    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToCsv = $OpenFileDialog.filename

    Clear-Host
    try{
        Write-Log "Importing CSV file"
        Write-Host "Importing CSV file"
        $csvImport = Import-CSV $pathToCsv

        $row = 1
        $vars.AllUsers.WithNoDevice = @{}
        foreach($user in $csvImport){

            $row++
            if([string]::IsNullOrEmpty($user.UserId) -or [string]::IsNullOrEmpty($user.UserPrincipalName)){
                Write-Log "Error on row $row. Skipping row."
                Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
                if([string]::IsNullOrEmpty($user.UserId)){
                    Write-Log "UserId is empty"
                    Write-Host "UserId is empty"
                }

                if([string]::IsNullOrEmpty($user.UserPrincipalName)){
                    Write-Log "UserPrincipalName is empty"
                    Write-Host "UserPrincipalName is empty"
                }

                ''
            }else{
                
                $vars.AllUsers.WithNoDevice.add($user.UserId, @{
                    userId = $user.UserId
                    userprincipalname = $user.UserPrincipalName
                })
            }
        }

        Clear-Host
        Write-Log "Sucecssfully imported CSV file: $($OpenFileDialog.filename)"
        Write-Host "Sucecssfully imported CSV file: $($OpenFileDialog.filename)" -ForegroundColor Magenta
    }catch{
        WRite-Log "An Error occurred reading CSV file"
        Write-Host "An Error occurred reading CSV file" -ForegroundColor Red
        Write-Error $_
    }

    #Users with a device
    Write-Log "Prompt user to select CSV file from Windows Explorer: All users with a device"
    Write-Host "Select CSV file from Windows Explorer: All users with a device"

    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToCsv = $OpenFileDialog.filename

    Clear-Host
    try{
        Write-Log "Importing CSV file"
        Write-Host "Importing CSV file"
        $csvImport = Import-CSV $pathToCsv

        $row = 1
        $vars.AllUsers.WithADevice = @{}
        foreach($user in $csvImport){
            $row++
            if([string]::IsNullOrEmpty($user.UserPrincipalName) -or 
                [string]::IsNullOrEmpty($user.approximateLastSignInDateTime) -or 
                [string]::IsNullOrEmpty($user.createdDateTime) -or 
                [string]::IsNullOrEmpty($user.deviceId) -or 
                [string]::IsNullOrEmpty($user.registrationDateTime) -or 
                [string]::IsNullOrEmpty($user.displayName)){

                Write-Log "Error on row $row. Skipping row."
                Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
                if([string]::IsNullOrEmpty($user.UserPrincipalName)){
                    Write-Log "UserPrincipalName is empty"
                    Write-Host "UserPrincipalName is empty"
                }

                if([string]::IsNullOrEmpty($user.approximateLastSignInDateTime)){
                    Write-Log "approximateLastSignInDateTime is empty"
                    Write-Host "approximateLastSignInDateTime is empty"
                }

                if([string]::IsNullOrEmpty($user.createdDateTime)){
                    Write-Log "createdDateTime is empty"
                    Write-Host "createdDateTime is empty"
                }

                if([string]::IsNullOrEmpty($user.deviceId)){
                    Write-Log "deviceId is empty"
                    Write-Host "deviceId is empty"
                }

                if([string]::IsNullOrEmpty($user.registrationDateTime)){
                    Write-Log "registrationDateTime is empty"
                    Write-Host "registrationDateTime is empty"
                }

                if([string]::IsNullOrEmpty($user.displayName)){
                    Write-Log "displayName is empty"
                    Write-Host "displayName is empty"
                }

                ''
            }else{
                
                $vars.AllUsers.WithADevice.add($user.deviceId, @{
                    UserPrincipalName = $user.UserPrincipalName
                    approximateLastSignInDateTime = $user.approximateLastSignInDateTime
                    createdDateTime = $user.createdDateTime
                    deviceId = $user.deviceId
                    registrationDateTime = $user.registrationDateTime
                    displayName = $user.displayName
                })
            }
        }

        Clear-Host
        Write-Log "Sucecssfully imported CSV file: $($OpenFileDialog.filename)"
        Write-Host "Sucecssfully imported CSV file: $($OpenFileDialog.filename)" -ForegroundColor Magenta
    }catch{
        WRite-Log "An Error occurred reading CSV file"
        Write-Host "An Error occurred reading CSV file" -ForegroundColor Red
        Write-Error $_
    }
}

Function Step7{
    #User with no device
    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    Write-Log "Prompt user to select JSON file from Windows Explorer: All users with no device"
    Write-Host "Select JSON file from Windows Explorer: All users with no device"

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "JSON (*.json)| *.json"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToJson = $OpenFileDialog.filename

    Clear-Host
    Write-Log "Importing JSON file"
    Write-Host "Importing JSON file"

    try {
        $errorImportingDevice = 0
        $jsonImport = Get-Content $pathToJson | ConvertFrom-Json 

        $vars.AllUsers.WithNoDevice = @{}
        foreach ($user in $jsonImport.PSObject.Properties.Value) {

            if ([string]::IsNullOrEmpty($user.UserId) -or [string]::IsNullOrEmpty($user.UserPrincipalName)) {
                Write-Log "Error on row $row. Skipping row."
                Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
                if ([string]::IsNullOrEmpty($user.UserId)) {
                    Write-Log "UserId is empty"
                    Write-Host "UserId is empty"
                }

                if ([string]::IsNullOrEmpty($user.UserPrincipalName)) {
                    Write-Log "UserPrincipalName is empty"
                    Write-Host "UserPrincipalName is empty"
                }

                ''
            }
            else {

                $vars.AllUsers.WithNoDevice.add($user.UserId, @{
                        userId            = $user.UserId
                        userprincipalname = $user.UserPrincipalName
                    })
            }
        }

        if ($errorImportingDevice -gt 0) {
            Write-Log "There is missing data in some of the json values. Please fix source data and re-run"
            Write-Host "There is missing data in some of the json values. Please fix source data and re-run" -ForegroundColor Red
            $vars.DevicesToCheck = @{}
        }
        else {
            Clear-Host
            Write-Log "Sucecssfully imported JSON file."
            Write-Host "Sucecssfully imported JSON file." -ForegroundColor Magenta
        }
    }
    catch {
        Write-Log "An Error occurred reading JSON file" 
        Write-Host "An Error occurred reading JSON file" -ForegroundColor Red
        Write-Error $_
    }

    #User with a device
    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    Write-Log "Prompt user to select JSON file from Windows Explorer: All users with a device"
    Write-Host "Select JSON file from Windows Explorer: All users with a device"

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "JSON (*.json)| *.json"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToJson = $OpenFileDialog.filename

    Clear-Host
    Write-Log "Importing JSON file"
    Write-Host "Importing JSON file"

    try {
        $errorImportingDevice = 0
        $jsonImport = Get-Content $pathToJson | ConvertFrom-Json 

        $vars.AllUsers.WithADevice = @{}
        foreach ($user in $jsonImport.PSObject.Properties.Value) {

            if ([string]::IsNullOrEmpty($user.UserPrincipalName) -or 
                [string]::IsNullOrEmpty($user.approximateLastSignInDateTime) -or 
                [string]::IsNullOrEmpty($user.createdDateTime) -or 
                [string]::IsNullOrEmpty($user.deviceId) -or 
                [string]::IsNullOrEmpty($user.registrationDateTime)) {

                Write-Log "Error on row $row. Skipping row."
                Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
                if ([string]::IsNullOrEmpty($user.UserPrincipalName)) {
                    Write-Log "UserPrincipalName is empty"
                    Write-Host "UserPrincipalName is empty"
                }

                if ([string]::IsNullOrEmpty($user.approximateLastSignInDateTime)) {
                    Write-Log "approximateLastSignInDateTime is empty"
                    Write-Host "approximateLastSignInDateTime is empty"
                }

                if ([string]::IsNullOrEmpty($user.createdDateTime)) {
                    Write-Log "createdDateTime is empty"
                    Write-Host "createdDateTime is empty"
                }

                if ([string]::IsNullOrEmpty($user.deviceId)) {
                    Write-Log "deviceId is empty"
                    Write-Host "deviceId is empty"
                }

                if ([string]::IsNullOrEmpty($user.registrationDateTime)) {
                    Write-Log "registrationDateTime is empty"
                    Write-Host "registrationDateTime is empty"
                }

                ''
            }
            else {
                
                $vars.AllUsers.WithADevice.add($user.deviceId, @{
                        UserPrincipalName             = $user.UserPrincipalName
                        approximateLastSignInDateTime = $user.approximateLastSignInDateTime
                        createdDateTime               = $user.createdDateTime
                        deviceId                      = $user.deviceId
                        registrationDateTime          = $user.registrationDateTime
                        displayName                   = $user.displayName
                    })
            }
        }

        if ($errorImportingDevice -gt 0) {
            Write-Log "There is missing data in some of the json values. Please fix source data and re-run"
            Write-Host "There is missing data in some of the json values. Please fix source data and re-run" -ForegroundColor Red
            $vars.DevicesToCheck = @{}
        }
        else {
            Clear-Host
            Write-Log "Sucecssfully imported JSON file."
            Write-Host "Sucecssfully imported JSON file." -ForegroundColor Magenta
        }
    }
    catch {
        Write-Log "An Error occurred reading JSON file" 
        Write-Host "An Error occurred reading JSON file" -ForegroundColor Red
        Write-Error $_
    }
}

Function Step8{
    #Users with no device
    Write-Log "Entering Step 8: Comparing devices and users"
    Write-Host "Comparing devices and users"

    foreach($user in $vars.AllUsers.WithADevice.GetEnumerator()){
        if(![string]::IsNullOrEmpty($vars.DevicesToCheck[$user.value.displayName][$user.Value.deviceId])){
            $vars.DevicesToCheck[$user.value.displayName][$user.Value.deviceId].userPrincipalName = $user.Value.UserPrincipalName
            $vars.DevicesToCheck[$user.value.displayName][$user.Value.deviceId].deleteStatus = "Do Not Delete"

        }
    }

    Write-Host "Complete. Re-run 1a or 1b to export the results"
    Write-Log "Complete. Re-run 1a or 1b to export the results"

}

Function menu{
    #Device check
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        $forgroundColour = "DarkGray"
        $numDevices = 0
    }else{
        $forgroundColour = "Green"
        $numDevices = 0
        foreach($device in $vars.DevicesToCheck.GetEnumerator()){
            foreach($deviceDuplicate in $device.value){
                foreach($identifiedDevice in $deviceDuplicate.GetEnumerator()){
                    $numDevices++
                }
            }
        }
    }

    #User check
    if(([string]::IsNullOrEmpty($vars.AllUsers.WithNoDevice) -or ($vars.AllUsers.WithNoDevice.count -eq 0) ) -and ([string]::IsNullOrEmpty($vars.AllUsers.WithADevice) -or ($vars.AllUsers.WithADevice.count -eq 0))){
        $userForgroundColour = "DarkGray"
    }else{
        $userForgroundColour = "Green"
    }

    if($userForgroundColour -notlike "Green" -and $forgroundColour -notlike "Green"){
        $step8ForgroundColour = "DarkGray"
    }else{
        $step8ForgroundColour = "Green"
    }

    '========================================================'
    Write-Host '          Azure AD Stale Device Clean UP Tool      ' -ForegroundColor Green 
    '========================================================'
    ''
    Write-Host "Please provide any feedback, comment or suggestion" -ForegroundColor Yellow
    ''
    Write-Host "$($vars.DevicesToCheck.count) unique devices in memory" -ForegroundColor Magenta
    Write-Host "$numDevices total devices in memory" -ForegroundColor Magenta
    Write-Host "$($vars.AllUsers.FromGraph.count) total users in memory from MS Graph" -ForegroundColor Magenta
    Write-Host "    $($vars.AllUsers.WithNoDevice.count) total users with no device in memory" -ForegroundColor Magenta
    Write-Host "    $($vars.AllUsers.WithADevice.count) devices with an associated user in memory" -ForegroundColor Magenta
    ''
    Write-Host "Enter (1) to get all Azure AD devices" -ForegroundColor Green
    ''
    Write-Host "    Enter (1a) to save devices in memory to csv" -ForegroundColor $forgroundColour
    ''
    Write-Host "    Enter (1b) to save devices in memory to json" -ForegroundColor $forgroundColour
    ''
    Write-Host "Enter (2) to import devices from csv file" -ForegroundColor Green
    ''
    Write-Host "Enter (3) to import devices from json file" -ForegroundColor Green
    ''
    Write-Host "Enter (4) to delete devices" -ForegroundColor $forgroundColour
    ''
    Write-Host "Enter (5) to get all Azure AD users and their devices" -ForegroundColor Green
    ''
    Write-Host "    Enter (5a) to save users and their devices in memory to csv" -ForegroundColor $userForgroundColour
    ''
    Write-Host "    Enter (5b) to save users and their devices in memory to json" -ForegroundColor $userForgroundColour
    ''
    Write-Host "Enter (6) to import users from csv file" -ForegroundColor Green
    ''
    Write-Host "Enter (7) to import users from json file" -ForegroundColor Green
    ''
    Write-Host "Enter (8) to compare users to devices" -ForegroundColor $step8ForgroundColour
    ''
    Write-Host "Enter (9) to Quit" -ForegroundColor Green
    ''

    $Selector =''
    $Selector = Read-Host -Prompt "Please make a selection, and press Enter" 

    While(($Selector -ne '1') -AND ($Selector -ne '1a') -AND ($Selector -ne '1b') -AND ($Selector -ne '2') -AND ($Selector -ne '3') -AND ($Selector -ne '4') -AND ($Selector -ne '5') -AND ($Selector -ne '5a') -AND ($Selector -ne '5b') -AND ($Selector -ne '6') -AND ($Selector -ne '7') -AND ($Selector -ne '8') -AND ($Selector -ne '9') -AND ($Selector -ne 'Debug')){

        $Selector = Read-Host -Prompt "Invalid input. Please make a correct selection from the above options, and press Enter" 
        
    }
    RunSelection -Selector $Selector
}

function RunSelection{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Selector
    )

    if($Selector -eq '1'){
        Clear-Host

        ''
        Write-Log "Menu Option: Get all Azure AD devices option has been chosen"
        Write-Host "Get all Azure AD devices option has been chosen" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step1
    }elseif($Selector -eq '1a'){
        Clear-Host

        ''
        Write-Log "Menu Option: Save output as csv option has been chosen"
        Write-Host "Save output as csv option has been chosen" -BackgroundColor Black
        ''
        Step1A
    }elseif($Selector -eq '1b'){
        Clear-Host

        ''
        Write-Log "Menu Option: Save output as json option has been chosen"
        Write-Host "Save output as json option has been chosen" -BackgroundColor Black
        ''
        Step1B
    }elseif($Selector -eq '2'){
        Clear-Host

        ''
        Write-Log "Menu Option: Import device from csv file option has been chosen"
        Write-Host "Import device from csv file option has been chosen" -BackgroundColor Black
        ''
        Step2
    }elseif($Selector -eq '3'){
        Clear-Host

        ''
        Write-Log "Menu Option: Import device from json file option has been chosen"
        Write-Host "Import device from json file option has been chosen" -BackgroundColor Black
        ''
        Step3
    }elseif($Selector -eq '4'){
        Clear-Host
        ''
        Write-Log "Menu Option: Delete devices option has been chosen"
        Write-Host "Delete devices option has been chosen" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step4
    }elseif($Selector -eq '5'){
        Clear-Host
        ''
        Write-Log "Menu Option: Get all Azure AD Users has been chosen"
        Write-Host "Get all Azure AD Users has been chosen" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step5
    }elseif($Selector -eq '5a'){
        Clear-Host
        ''
        Write-Log "Menu Option: Save users and their devices in memory to csv"
        Write-Host "Save users and their devices in memory to csv" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step5A
    }elseif($Selector -eq '5b'){
        Clear-Host
        ''
        Write-Log "Menu Option: Save users and their devices in memory to json"
        Write-Host "Save users and their devices in memory to json" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step5B
    }elseif($Selector -eq '6'){
        Clear-Host
        ''
        Write-Log "Menu Option: Import users from csv file option has been chosen"
        Write-Host "Import users from csv file option has been chosen" -BackgroundColor Black
        ''
        Step6
    }elseif($Selector -eq '7'){
        Clear-Host
        ''
        Write-Log "Menu Option: Import users from json file option has been chosen"
        Write-Host "Import users from json file option has been chosen" -BackgroundColor Black
        ''
        Step7
    }elseif($Selector -eq '8'){
        Clear-Host
        ''
        Write-Log "Menu Option: Compare users to existing devices"
        Write-Host "Compare users to existing devices" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step8
    }elseif($Selector -eq 'Debug'){
        Write-Log "Menu Option: Debug option chosen"
        $vars.DelegateToken | ConvertTo-Json
    }elseif($Selector -eq '9'){
        break
    }

    ''
    $null = Read-Host 'Press Any Key or Enter to return to the menu'
    ''
    cls
    menu
}

cls
menu