# Import the data from CSV file and assign it to variable
$ADUsers = Import-Csv "C:\Users\helpdesk.NIGHTCITY\Downloads\CreateBulkADUsers.csv"

# Specify target Security Group to add
$AddGroup = ""

#Specify target Security Group to Remove
$RemoveGroup = ""

ForEach ($User in $ADUsers) {
    # Strip @domain from SamAccountName if present
    $SAM = $User.SamAccountName -replace "@.*$", ""

    # Retrieve AD User object
    $ADUser = Get-ADUser -Filter "SamAccountName -eq '$SAM'"

    Write-Host "Disabling and Adding AD User to Group: $($ADUser.SamAccountName)"

    # Disable user
    Disable-ADAccount -Identity $ADUser

    # Add user to target Security Group
    Add-ADGroupMember -Identity $AddGroup -Members $ADUser

    # Remove users from target Security Group
    Remove-ADGroupMember -Identity $RemoveGroup -Members $ADUser

    # End Message
    Write-Host "$($ADUser.SamAccountName) has been Disabled and Added to '$AddGroup'"
    Write-Host "$($ADUser.SamAccountName) has been removed from '$RemoveGroup'" 
}

Write-Host "Completed Script"
