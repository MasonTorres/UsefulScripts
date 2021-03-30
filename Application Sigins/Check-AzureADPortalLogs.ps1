# Export the new Sign-in logs from Azure AD: ApplicationSignIns, InteractiveSignIns, MSISignIns, NonInteractiveSignIns
# Add the path to the CSV files in the below variables
# Run script

$vars = @{
    # Azure AD CSV logs location
    Import = @{
        ApplicationSignIns = "ApplicationSignIns_2021-02-28_2021-03-30.csv"
        InteractiveSignIns = "InteractiveSignIns_2021-02-28_2021-03-30.csv"
        MSISignIns = "MSISignIns_2021-02-28_2021-03-30.csv"
        NonInteractiveSignIns = "NonInteractiveSignIns_2021-02-28_2021-03-30.csv"
    }
    # CSV Export location for the audit
    Export = @{
        Location = "Sign-InEvents-AppSP.csv"
    }
    ScriptStartTime = Get-Date
}

# Add each event to a hash table, loop over csv logs adding applications or service principals as they appear, counting the number of entries.
$SignInEvents = @{}
foreach($log in $vars.Import.GetEnumerator()){
    $csvLog = Import-Csv $vars.Import.($log.Name)

    foreach($logEntry in $csvLog){
        if($null -eq $SignInEvents.($logEntry.'Application ID ')){
            $details = @{
                AppName = $logEntry.Application
                ServicePrincipalName = $logEntry.'Service principal name'
                ApplicationID = $logEntry.'Application ID '
                SignInType = $log.Name
                Count = 1
            }

            $SignInEvents.($logEntry.'Application ID ') = $details
        }else{
            $SignInTypes = ""
            if($SignInEvents.($logEntry.'Application ID ').SignInType -notlike "*$($log.Name)*"){
                $currentName = $SignInEvents.($logEntry.'Application ID ').SignInType
                $SignInTypes =  "$currentName,$($log.Name)"
            }else{
                $SignInTypes = $SignInEvents.($logEntry.'Application ID ').SignInType 
            }
            $details = @{
                AppName = $logEntry.Application
                ServicePrincipalName = $logEntry.'Service principal name'
                ApplicationID = $logEntry.'Application ID '
                SignInType = $SignInTypes
                Count = ($SignInEvents.($logEntry.'Application ID ').count + 1)
            }
            $SignInEvents.($logEntry.'Application ID ') = $details
        }
    }
}

# Flatten the hashtable and export to CSV
$output = @()

foreach($app in $SignInEvents.GetEnumerator()){
    $details = [ordered]@{}
    $details.add("App Name", $SignInEvents.($app.Name).AppName)
    $details.add("Service principal name", $SignInEvents.($app.Name).ServicePrincipalName)
    $details.add("App ID", $app.Name)
    $details.add("SignInTypes", $SignInEvents.($app.Name).SignInType)
    $details.add("Count", $SignInEvents.($app.Name).Count)
    $output += New-Object PSObject -Property $details
}

$output | export-csv -NoTypeInformation $vars.Export.Location