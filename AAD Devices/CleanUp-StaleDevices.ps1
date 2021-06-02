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
         [string] $Body
    )

    $ReturnValue = $null
    $OneSuccessfulFetch = $null
    $PermissionCheck = $false
    $RetryCount = 0
    $ResultNextLink = $Uri
    $batchCount = 0

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
            $batchCount++
            Write-Host -NoNewline "`rRecevied pages from Microsoft Graph API $batchCount"
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
                ##$Token.AccessToken = Get-AppToken -tenantId $Token.TenantID -clientId $Token.ClientID -clientSecret $Token.ClientSecret
                try{
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
                    $vars.ScriptStartTime = Get-Date
                    
                }catch{
                    $e = $_
                    $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }

                $OneSuccessfulFetch = $false
            }
            elseif ($statusCode -eq 401 -and $PermissionCheck -eq $false) {
                # In the case we are making multiple individual calls to Invoke-MSGraph we may need to check the access token has expired in between calls.
                # i.e the check above never occurs if MS Graph returns only one page of results.
                Write-Output "Retrying..."
                ##$Token.AccessToken = Get-AppToken -tenantId $Token.TenantID -clientId $Token.ClientID -clientSecret $Token.ClientSecret
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
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
        $vars.ScriptStartTime = Get-Date
        
    }catch{
        $e = $_
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    }

    Clear-Host
}

Function Step1{
    # Get our Access Token for Microsoft Graph API
    #$vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    try{
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret -refreshToken $vars.DelegateToken.RefreshToken
        $vars.ScriptStartTime = Get-Date
        
    }catch{
        $e = $_
        $vars.DelegateToken = Get-DelegateToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
    }

    ''
    Write-Host "Getting devices..."

    # Get all risky users from Microsoft Graph API
    $vars.AllDevices = Get-Devices -Token $vars.DelegateToken

    # Create a Hashtable with unique device names, then add any duplicate devices as child objects.
    $vars.DevicesToCheck = @{}
    foreach($device in $vars.AllDevices){
        
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

    ''
    Write-Host "Number of Devices: $($vars.AllDevices.count)" -ForegroundColor Magenta
    Write-Host "Number of Unique devices by name: $($vars.DevicesToCheck.count)" -ForegroundColor Magenta
    ''
}

Function Step2{
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        Write-Host "No devices found. Please run option 1" -ForegroundColor Yellow
    }else{
        # Export the hashtable as a CSV file
        $CsvOutput = @()
        foreach($device in $vars.DevicesToCheck.GetEnumerator()){
            
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
    }
}

Function Step3{
    if([string]::IsNullOrEmpty($vars.DevicesToCheck) -or $vars.DevicesToCheck.count -eq 0){
        Write-Host "No devices found. Please run option 1" -ForegroundColor Yellow
    }else{
        # Export the hashtable as a JSON object
        $jsonOutput =  $vars.DevicesToCheck | ConvertTo-Json
        $jsonOutput | Out-File DuplicateDevices.json
    }
}

Function Step4{
    
    Write-Host "Select CSV file from Windows Explorer"

    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToCsv = $OpenFileDialog.filename

    Clear-Host

    $csvImport = Import-CSV $pathToCsv

    $row = 1
    $vars.DevicesToCheck = @{}
    foreach($device in $csvImport){

        $row++
        if([string]::IsNullOrEmpty($device.Computer) -or [string]::IsNullOrEmpty($device.CreatedDateTime) -or [string]::IsNullOrEmpty($device.ObjectID) -or [string]::IsNullOrEmpty($device.DeviceID)){
            Write-Host "Error on row $row. Skipping row." -ForegroundColor Red
            if([string]::IsNullOrEmpty($device.Computer)){
                Write-Host "Computer is empty"
            }

            if([string]::IsNullOrEmpty($device.CreatedDateTime)){
                Write-Host "CreatedDateTime is empty"
            }

            if([string]::IsNullOrEmpty($device.ObjectID)){
                Write-Host "ObjectID is empty"
            }

            if([string]::IsNullOrEmpty($device.DeviceID)){
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
                        displayName = $device.displayName
                        deleteStatus = $device.DeleteStatus
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
                    }
                })
            }
        }
    }
}

Function Step5{
    #Get file from Windows Explorer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    Write-Host "Select JSON file from Windows Explorer"

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "JSON (*.json)| *.json"
    $OpenFileDialog.ShowDialog() | Out-Null
    $pathToJson = $OpenFileDialog.filename

    Clear-Host

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
                    Write-Host "$($dupDevice.objectId): Computer is empty"
                }
    
                if([string]::IsNullOrEmpty($dupDevice.createdDateTime)){
                    Write-Host "$($dupDevice.displayName): CreatedDateTime is empty"
                }
    
                if([string]::IsNullOrEmpty($dupDevice.objectId)){
                    Write-Host "$($dupDevice.displayName): ObjectID is empty"
                }
    
                if([string]::IsNullOrEmpty($dupDevice.deviceId)){
                    Write-Host "$($dupDevice.displayName): DeviceID is empty"
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
                        }
                    })
                }
            }
        }

        $vars.DevicesToCheck += $device
    }
}

Function Step6{

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
                    $deleteResult = Invoke-DeleteDevicesBatch -Token $vars.DelegateToken -batchRequest $batchRequest
                    
                }
            }
        }
    }
}

Function menu{
    
    '========================================================'
    Write-Host '        Device Deletion Tool          ' -ForegroundColor Green 
    '========================================================'
    ''
    Write-Host "Please provide any feedback, comment or suggestion" -ForegroundColor Yellow
    ''
    Write-Host "Enter (1) to get all Azure AD devices" -ForegroundColor Green
    ''
    Write-Host "Enter (2) to save output as csv" -ForegroundColor Green
    ''
    Write-Host "Enter (3) to save output as json" -ForegroundColor Green
    ''
    Write-Host "Enter (4) to import devices from csv file" -ForegroundColor Green
    ''
    Write-Host "Enter (5) to import devices from json file" -ForegroundColor Green
    ''
    Write-Host "Enter (6) to delete devices" -ForegroundColor Green
    ''
    Write-Host "Enter (9) to Quit" -ForegroundColor Green
    ''

    $Num =''
    $Num = Read-Host -Prompt "Please make a selection, and press Enter" 

    While(($Num -ne '1') -AND ($Num -ne '2') -AND ($Num -ne '3') -AND ($Num -ne '4') -AND ($Num -ne '5') -AND ($Num -ne '6') -AND ($Num -ne '7') -AND ($Num -ne '9')){

        $Num = Read-Host -Prompt "Invalid input. Please make a correct selection from the above options, and press Enter" 
        
    }
    RunSelection -Num $Num
}

function RunSelection{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [int] $Num
    )

    if($Num -eq '1'){
        Clear-Host

        ''
        Write-Host "Get all Azure AD devices option has been chosen" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step1
    }elseif($Num -eq '2'){
        Clear-Host

        ''
        Write-Host "Save output as csv option has been chosen" -BackgroundColor Black
        ''
        Step2
    }elseif($Num -eq '3'){
        Clear-Host

        ''
        Write-Host "Save output as json option has been chosen" -BackgroundColor Black
        ''
        Step3
    }elseif($Num -eq '4'){
        Clear-Host

        ''
        Write-Host "Import device from csv file option has been chosen" -BackgroundColor Black
        ''
        Step4
    }elseif($Num -eq '5'){
        Clear-Host

        ''
        Write-Host "Import device from json file option has been chosen" -BackgroundColor Black
        ''
        Step5
    }elseif($Num -eq '6'){
        Clear-Host
        ''
        Write-Host "Delete devices option has been chosen" -BackgroundColor Black
        ''
        Invoke-CheckVariables
        Step6
    }elseif($Num -eq '9'){
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