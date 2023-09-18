$userID = "272b8ba7-22db-4ddf-94c3-ad9416b7394e"

$user = get-azureaduser -ObjectId $userID
$memberships = Get-AzureADUserMembership -ObjectId $userID -All:$true
function CheckGroup {
    param(
        [parameter(Mandatory, Position = 0)]
        $GroupID
    )

    $groupMembers = Get-AzureADGroupMember -ObjectId $GroupID

    $userInGroup = $false
    foreach ($member in $groupMembers) {
        if ($member.ObjectId -like $userID) {
            $userInGroup = $true
            Write-Host "User is direct in group $($GroupID)" -ForegroundColor Green
        }

        if ($member.ObjectType -eq "Group") {
            $userInGroup = CheckGroup -GroupID $member.ObjectId
            if ($userInGroup -eq $true) {
                Write-Host "User in subgroup" -ForegroundColor Yellow
            }
        }
    }
    return $userInGroup
}

foreach ($group in $memberships) {
    CheckGroup -GroupID $group.ObjectID
}