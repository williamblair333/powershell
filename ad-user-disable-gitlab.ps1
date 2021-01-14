<# 
File:       ad-user-disable-gitlab.ps1
Date:       2021JAN14
Author:     William Blair
Contact:    williamblair333@gmail.com
#>

<#
This script will do the following:
- strip all AD groups
- add users to a group called 'Accounts-Disabled'
- reassign new primary group
- remove all other ad groups
- disable the account
- set the account to expire
- move the account to another ou
#>

Import-Module -Name ActiveDirectory

$date = (get-date (get-date).addDays(-1) -Format "MM/dd/yyyy")

# Prompt for user
$user = Read-Host -Prompt 'Enter AD account username'

# A group we're gonna add
$group = get-adgroup "Accounts-Disabled" -properties @("primaryGroupToken")

# Add the disabled group
Add-ADGroupMember -Identity $group -Members $user

# Set the primary group
Get-ADUser $user | set-aduser -replace @{primaryGroupID=$group.primaryGroupToken}

# Remove all the groups except primary group
Get-ADUser -Identity $user -Properties MemberOf | ForEach-Object {
  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
}

#Disable the Account
Disable-ADAccount -Identity $user

# Set the Account to expire
Set-ADAccountExpiration -Identity $user -DateTime $date

# Move the account to target OU
Get-ADUser $user | Move-ADObject -TargetPath 'OU=Disabled,OU=Fire,DC=waynetwp,DC=local'