# o365-account-user-query.ps1
# William Blair
# williamblair333@gmail.com

# This script will get information about a particular email account. 
# You could run the script like this:  powershell o365-account-user-query.ps1 joe.blow@contoso.com

param($Email)

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Get-PSSession | Remove-PSSession

$365Logon = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $365Logon -Authentication Basic -AllowRedirection

Import-PSSession $Session

Get-Recipient | where {$_.EmailAddresses -match $Email} | fL Name, RecipientType,emailaddresses

Get-PSSession | Remove-PSSession