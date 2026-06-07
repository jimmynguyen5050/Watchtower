Import-Module ActiveDirectory

$Users = Import-Csv "C:\Path_To_CSV"

foreach ($User in $Users) {

    # Strip domain from SamAccountName column to get just the logon name
    $Username = $User.SamAccountName.Split("@")[0]

    # Check if user already exists - use string filter, not scriptblock
    if (Get-ADUser -Filter "SamAccountName -eq '$Username'") {
        Write-Warning "User $Username already exists in Active Directory. Skipping."
    }
    else {
        $userCreation = @{
            SamAccountName        = $Username
            GivenName             = $User.givenName
            Surname               = $User.sn
            Name                  = $User.Name
            UserPrincipalName     = $User.SamAccountName
            Company               = $User.Company
            AccountPassword       = (ConvertTo-SecureString $User.Password -AsPlainText -Force)
            Enabled               = $true
            ChangePasswordAtLogon = $false
            PasswordNeverExpires  = $true
        }

        New-ADUser @userCreation
        Write-Host "Created user: $Username" -ForegroundColor Cyan
    }
}
