import-module ActiveDirectory

get-command -Module ActiveDirectory

$Pass = ConvertTo-SecureString "Password" -AsPlainText -Force 

Set-ADAccountPassword -Identity username -NewPassword $Pass -Reset 
Set-ADUser -Identity username -ChangePasswordAtLogon $true
