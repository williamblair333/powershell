# o365-users-licenses-add.ps1
# William Blair
# Support@waynetwp.org

# This script will assign o365 licenses to users listed in the Email.csv file
# This is all the license types to choose from.  Obviously you will need free licenses
# Reference here:  https://docs.microsoft.com/en-us/office365/enterprise/powershell/assign-licenses-to-user-accounts-with-office-365-powershell
# And here: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell#connect-with-the-azure-active-directory-powershell-for-graph-module

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

#Connect-AzureAD

# Cycle through the script pulling users from the email.csv file
Import-CSV "C:\temp\Email.csv" | ForEach-Object {

$userUPN=$userUPN=$_.member
$planName="TEAMS_EXPLORATORY"
$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
$LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$LicensesToAssign.AddLicenses = $License
Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $LicensesToAssign

}