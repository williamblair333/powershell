# o365-group-distro-users-get.ps1
# William Blair
# williamblair333@gmail.com

# This script will get users from Email.csv file to the specified distro group

# Supply the Distro Group and a file location (e.g C:\temp
param($DGName, $FileLocation)

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Get-PSSession | Remove-PSSession

$365Logon = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $365Logon -Authentication Basic -AllowRedirection

Import-PSSession $Session

$Filename = $FileLocation + "\" + $DGName + ".csv" 

Get-DistributionGroupMember -Identity $DGName -ResultSize Unlimited | 
	Select Name, PrimarySMTPAddress, RecipientType |
		Export-CSV  $Filename -NoTypeInformation -Encoding UTF8

Get-PSSession | Remove-PSSession