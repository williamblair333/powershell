# o365-users-licenses-query-export.ps1
# William Blair
# williamblair333@gmail.com

# This script will query your org for specified license pool showing what license is assigned to who.

Connect-MsolService
Get-MsolAccountSku

#Generic text dump export

<# SkuPartNumber
-------------
STREAM
WINDOWS_STORE
ENTERPRISEPACK
FLOW_FREE
RIGHTSMANAGEMENT
EXCHANGESTANDARD
TEAMS_EXPLORATORY
STANDARDPACK
-------------
#>

#gets all users with license "EXCHANGESTANDARD", and returns only fields after -Property in neat csv format
Get-MsolUser | Where-Object {($_.licenses).AccountSkuId -match "EXCHANGESTANDARD"} `
| Select-Object -Property LastName, FirstName, DisplayName, UserPrincipalName, Office, PhoneNumber, Fax, WhenCreated, LastPasswordChangeTimestampitle, IsLicensed `
| Export-Csv -path .\o365__lic_users.csv -NoTypeInformation