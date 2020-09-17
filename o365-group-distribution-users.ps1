<# o365-group-distribution-users.ps1
 William Blair
 williamblair333@gmail.com
 Housekeeping script for o365 distribution groups without params
#>

# Get credentials and authenticate
$365Logon = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
-Credential $365Logon -Authentication Basic -AllowRedirection

Import-PSSession $Session

# Clears the screen
Clear-Host

# Loop to check if the distribution group is valid.
Do {
	$GroupName = Read-Host -Prompt "Input Group Name"

	<# Since the error isn't local, try / catch wouldn't work.  Most errors here are probably
	 missing / wrong group name, so catch all and loop until they get it right or quit
	#>
	$ExistingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction 'SilentlyContinue'

	# null out the variable and keep trying
	if(-not $existingGroup) {
		Write-Host "$GroupName object not found. Did you spell it correctly?  Try again or CTRL-C to quit." 
		$GroupName = $null
	}
	
	else {
		Write-Host ""
		Write-Host "Group name exists, continuing"
		Write-Host ""
	}

	}
While ($GroupName -eq $null)

	# There are probably better ways to do menus, but this works
Do {

	Write-Host ""
	Write-Host "What would you like to do?"
	Write-Host ""
	Write-Host "	(l) List Group Member(s)"
	Write-Host "	(af) Add a Group Member(s) from the csv file"
	Write-Host "	(rf) Remove a Group Member(s) from the csv file"
	Write-Host "	(ap) Add a Group Member from the prompts here"
	Write-Host "	(rp) Remove a Group Member from the prompts here"
	Write-Host "	(c) Clear the screen"
	Write-Host "	(q) Quit"
	Write-Host ""

	# The switch variable...
	$InputChoice = Read-Host -Prompt "Enter your choice"

	switch($InputChoice) {
		l	{Get-DistributionGroupMember -Identity $GroupName | Format-Table Name -Auto }
		
			<# Have a csv named below with email address of each user in it.  
			 Adds users from the csv to the distribution group
			'member' is the only listed value 
			#>
	    af	{Import-CSV "$PSScriptRoot\o365-group-distribution-users-add.csv" |  `
		     ForEach-Object { Add-DistributionGroupMember –Identity $GroupName –Member $_.member} }

			<# Same as the af section.  Only this will remove users.  Easier to
			 use two separate csv files and dummy proof for some non-tech users.  
			'member' is the only listed value 
			#>
	    rf	{Import-CSV "$PSScriptRoot\o365-group-distribution-users-remove.csv" |  `
		     ForEach-Object { Remove-DistributionGroupMember –Identity $GroupName –Member $_.member} }

		ap	{$UserName = $null
		     $UserName = Read-Host -Prompt "Input User's Display Name"
		     Add-DistributionGroupMember -Identity $GroupName -Member $UserName
		     Get-DistributionGroupMember -Identity $GroupName}		

		rp	{$UserName = Read-Host -Prompt "Input User's Display Name"
		     Remove-DistributionGroupMember -Identity $GroupName -Member $UserName
		     Get-DistributionGroupMember -Identity $GroupName }

		c 	{Clear-Host}

			# Disconnect from o365 and exit the program
		q 	{Remove-PSSession $Session
		     Break}
	}

	}
While ($InputChoice -ne 'q')