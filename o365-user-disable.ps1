<# 
File:       o365-user-disable.ps1
Date:       2021JAN15
Author:     William Blair
Contact:    williamblair333@gmail.com
Note:       run all files from c:\temp
#>

<#
This script will do the following:
- change user email address
- remove all groups
- add group called 'disabled'
- remove all licenses
#>

# Get credentials and authenticate
$365Logon = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $365Logon -Authentication Basic -AllowRedirection

Import-PSSession $Session

# Clears the screen
Clear-Host

# Prompt for user
$upn = Read-Host -Prompt 'Enter the email address of the employee you wish to change'

# string called disabled which we will use later concactenate with email address
$string_disabled = 'disabled.'

# This is going to be the new disabled email account name
$upn_disabled = $string_disabled + $upn

# Change the email account name
Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName $upn_disabled

#Let's update $upn now
$upn = $upn_disabled

#Remove user from all assigned groups
# "C:\temp\Remove_User_All_Groups.ps1" -Identity $upn -IncludeAADSecurityGroups -IncludeOffice365Groups
& '.\Remove_User_All_Groups.ps1' -Identity $upn -IncludeAADSecurityGroups -IncludeOffice365Groups

#Remove all license from user acount
(get-MsolUser -UserPrincipalName $upn).licenses.AccountSkuId |
foreach{
    Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $_
}