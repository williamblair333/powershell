# o365-group-distro-users-add.ps1
# William Blair
# williamblair333@gmail.com

# This script will add users from Email.csv file to the specified distro group

$365Logon = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $365Logon -Authentication Basic -AllowRedirection

Import-PSSession $Session

Import-CSV "C:\temp\email.csv" | ForEach-Object {
Add-DistributionGroupMember –Identity "groupname" –Member $_.member
}

Remove-PSSession $Session