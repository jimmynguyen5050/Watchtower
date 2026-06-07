Import-Module ActiveDirectory

$splat = @{
    Name = 'Test3'
    AccountPassword = (Read-Host -AsSecureString 'Account Password')
    PasswordNeverExpires = $true
    Enabled = $true
}


New-ADUser @splat 

