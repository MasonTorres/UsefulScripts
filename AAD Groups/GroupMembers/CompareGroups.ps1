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
# Assign Application permissions: 	Group.Read.All https://docs.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0
# Create Application secret
# Set the App Registration to Public
# Update $vars variable below: ClientSecret, ClientID and TenantID 

# This script will delete devices from Azure AD.
# App ClientSecret, ClientID and TenantID are needed for delegate token refresh

param (
    [Parameter( ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true,
                HelpMessage="Get group membership count")]
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
    ScriptStartTime = Get-Date

    AllDevices = @{}
    DevicesToCheck = @{}

    LogFile = "Logs.txt"
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
        scope         = "https://graph.microsoft.com/.default offline_access"
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
        $headers.add("ConsistencyLevel", "Eventual");
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
            #For batches 
            if($result.responses){
                $ReturnValue = $result
            }
            #For $count
            if($result -match "^\d+$"){
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

                try{
                    $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }catch{
                    Write-Log "Could not get new client credential token."
                    Write-Log "Trying again."
                    $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }

                $OneSuccessfulFetch = $false
            }
            elseif ($statusCode -eq 401 -and $PermissionCheck -eq $false) {
                # In the case we are making multiple individual calls to Invoke-MSGraph we may need to check the access token has expired in between calls.
                # i.e the check above never occurs if MS Graph returns only one page of results.
                Write-Log "Retrying...Getting new access token"
                
                try{
                    $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
                }catch{
                    Write-Log "Could not get new client credential token."
                    Write-Log "Trying again."
                    $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
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

        #Get a new token after 45mins time has elapsed
        if(( Get-Date $vars.ScriptStartTime).AddMinutes(45) -lt (Get-Date)){
            Write-Log "45 mins elapsed. StartTime: Get-Date $($vars.ScriptStartTime)"

            try{
                $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
            }catch{
                Write-Log "Could not get new client credential token."
                Write-Log "Trying again."
                $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
            }
        }
    }

    return $ReturnValue
}

function Get-Groups{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token
    )

    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=onPremisesSyncEnabled eq true&`$top=999&`$select=displayName,id"
    $method = "GET"

    $Groups = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -DebugLots $true

    return $Groups
}

function Get-GroupsMemberCount{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [hashtable] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupId
    )

    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$count"
    $method = "GET"

    $Group = Invoke-MSGraph -Token $Token -Uri $uri -Method $method -DebugLots $true -useConsistencyLevel $true -ShowProgress $false

    return $Group
}


# Get our client credential access token
try{
    $vars.Token.AccessToken = Get-AppToken -tenantId $vars.Token.TenantID -clientId $vars.Token.ClientID -clientSecret $vars.Token.ClientSecret
}catch{
    $e = $_
    Write-Error $e
}

#Get all groups
$Groups = Get-Groups -Token $vars.token 

#Loop through groups and get their membership count
$output = @()
$groupProcessCount = 0
$AllGroupsCount = $Groups.count

Set-Content GroupsWithMembershipCount.csv "displayname,id,GroupMemberCount"

foreach($group in $Groups){
    $percent = [math]::Round($(($groupProcessCount / $AllGroupsCount) * 100),2)
    Write-Host -NoNewline "`rCreating batches of users $groupProcessCount / $AllGroupsCount $percent %"

    $groupCount = Get-GroupsMemberCount -Token $vars.Token -GroupId $group.id

    # $details = [ordered]@{}
    # $details.add("displayName",$group.displayName)
    # $details.add("id",$group.id)
    # $details.add("GroupMemberCount",$groupCount)
    # $output+= New-Object PSObject -Property $details
    
    ### Using Add-Content instead of an array to avoid high memory consumption when a large number of groups are processed
    Add-Content GroupsWithMembershipCount2.csv "$($group.displayName),$($group.id),$groupCount"
    $groupProcessCount++
}

# $output | Export-csv -NoTypeInformation GroupsWithMembershipCount.csv

