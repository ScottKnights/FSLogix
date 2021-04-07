# Migrate-profiles.ps1
# Written by Scott Knights
# V1.20210218.1 Initial Release 

<#
.SYNOPSIS
Migrates a list of user profiles to FSLogix.

.DESCRIPTION
Extracts list of users from a specified textfile or an AD OU and migrates their profile to an FSLogix VHD using FRX.EXE.
Once the VHD has beem migrated, the user will be made the owner and granted full control of the folder and VHD file.
Creates FSLogix folders in reversed <USERNAME><SID> format. Modify if you use a different format.
Run on the terminal server containing the local profiles.
Will not process the username if any of the following are true:
	User does not exist in AD
	User is logged on
	User does not have a profile on the server
	User already has a VHD file

Active Directory module for Windows (RSAT-AD-PowerShell) must be installed. Script will exit if it is not.

If an AD group is specified in the $ADGroup variable, user will be added to the group after successful migration. Allows for selective enablement of FSLogix for migrated profiles.

Occasionally FRX will return a non zero error code for no discernible reason and fail the migration. 
It will try twice as it often succeeds on the 2nd attempt. If still an issue after 2nd attempt, try running the script again before investigating further.

Modify the following variables to reflect your environment:
$profpath	UNC path to the share which will hold the FSLogix VHD files.
$userlist	Text file of samAccountnames of users you wish to process. If specified this will be used instead of getting users from AD.
$searchbase	DN of the OU you wish to process for usernames. If left blank and no userlist supplied, all users in AD will be processed.
$logfile	Full path of the log file.
$ADGroup	If specified, users with successfully migrated profiles will be added to this group
$vhdsize	VHD size in MB. VHD is dynamic so this sets the maximum file size.
$frxpath	Path to FRX.EXE. This should always be "c:\Program Files\FSLogix\Apps\frx.exe" if FSLogix is installed.
#>

#region Functions
# ============================================================================
# Functions
# ============================================================================
function process-username {

	# Get Username as mandatory parameter
	param (
		[Parameter(Mandatory=$true)]
		[string]$username
	)

	# Get SID of user. Return if user does not exist.
	try {
		$SID=(get-aduser -identity $username).sid.value
	} catch {
		add-content -Path $logfile "$username - Not found in AD"
		write-host "$username - Not found in AD"
		Return
	}

	# Return if user is logged on
	$ErrorActionPreference="Stop"
	try {
		query user $username 2>&1 |out-null
		add-content -Path $logfile "$username - Currently logged in"
		write-host "$username - Currently logged in"
		Return
	} catch {
		$ErrorActionPreference="Continue"
	}

	# Return if no profile path registry value
	$ErrorActionPreference="Stop"
	try {
		$profsrc=Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -Name ProfileImagePath
		$ErrorActionPreference="Continue"
	} catch {
		add-content -Path $logfile "$username - No profile path value in the registry"
		write-host "$username - No profile path value in the registry"
		return
	}

	# Return if the profile path folder does not exist
	if (-NOT (test-path $profsrc.ProfileImagePath)) {
		add-content -Path $logfile "$username - Profile path folder does not exist. Orphaned Profile"
		write-host "$username - Profile path folder does not exist. Orphaned Profile"
		Return
	}

	# Form folder and VHD names
	$fldname=$profpath+"\"+$username+"_"+$SID
	$vhdname="Profile_"+$username+".vhd"
	$fullpath=$fldname+"\"+$vhdname

	# Return if VHD already exists
	if (test-path $fullpath) {
		add-content -Path $logfile "$username - VHD already exists"
		write-host "$username - VHD already exists"
		Return
	}

	add-content -Path $logfile "$username - Processing"
	Write-host "$username - Processing"
	# Create Folder
	md $fldname -erroraction silentlycontinue |out-null

	# Set Folder Owner
	$usracct=New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $env:userdomain"\"$username
	$fldacl=get-acl $fldname
	$fldacl.setowner($usracct)
	set-acl -path $fldname -aclobject $fldacl

	# Migrate profile using FRX.EXE

	try {
#		& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" -verbose # Use for troubleshooting
		& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" 2>&1 |out-null
		$returncode=$lastexitcode
		if ($returncode -ne 0) {
			Write-host "$username - FRX return code did not equal 0. Trying again."
			& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" 2>&1 |out-null
			$returncode=$lastexitcode
			if ($returncode -ne 0) {
				add-content -Path $logfile "$username - FRX return code did not equal 0 - Investigate"
				Write-host "$username - FRX return code still did not equal 0"
				rd $fldname -erroraction silentlycontinue |out-null
				return
			}
		}
	} catch {
		add-content -Path $logfile "$username - FRX returned an error - Investigate"
		Write-host "$username - FRX returned an error"
		rd $fldname -erroraction silentlycontinue |out-null
		return
	}

	# Pause to give the system time to clean up. Without pause, FRX may error on next profile.
	start-sleep -seconds 5
	
	# Set VHD Owner and Permissions
	$vhdacl=get-acl $fullpath
	$vhdacl.setowner($usracct)
	$aclrights="FullControl"
	$acltype="Allow"
	$aclarglist=$usracct, $aclrights, $acltype
	$aclrule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $aclarglist
	$vhdacl.SetAccessRule($aclrule)
	set-acl -path $fullpath -aclobject $vhdacl

	# Add User to Group
	if ($adgroup) {
		try {	
			get-adgroup $adgroup -erroraction silentlycontinue |out-null
			Add-ADGroupMember -identity $adgroup -members $username
		} catch {
			write-host "$username - cannot add user to group $adgroup"
			add-content -Path $logfile "$username - cannot add user to group $adgroup"
		}
	}
}
#endregion Functions

#region Variables
# ============================================================================
# Variables
# ============================================================================

# Set destination folder path for FSLOGIX VHD files
$profpath="\\SERVER\SHARE"

# Set the file path of a list of usernames you want to process. If specified, users will not be queried from AD.
#$userlist=".\userlist.txt"

# Set the searchbase to specify the OU you want to process for usernames. Not used if $userlist is set.
#$searchbase="DC=domain,DC=com"

# Set log file path
$logfile=".\profmig.log"

# AD Group to add users to after migration
#$adgroup="RL-FSLogix Users"

# Set the VHD size in MB
$vhdsize=300000

# Set path to FRX.EXE. Should not need to change this
$frxpath="c:\Program Files\FSLogix\Apps\frx.exe"

#endregion Variables

#region Execute
# ============================================================================
# Execute
# ============================================================================
$date=get-date
clear-host
set-content -Path $logfile "start run at $date"
write-host "start run at $date"

# Test if FRX.EXE exists
if (-NOT (test-path $frxpath)) {
	write-host "FRX.EXE is missing"
	add-content -Path $logfile "FRX.EXE is missing. FSLogix is probably not installed. Exiting."
	Return
}

# Test if RSAT-AD-PowerShell is installed.
$rsatad=Get-WindowsFeature *RSAT-AD-PowerShell*
if (-not ($rsatad.installed)) {
	write-host "Active Directory module for Windows (RSAT-AD-PowerShell) is not installed."
	add-content -Path $logfile "Active Directory module for Windows (RSAT-AD-PowerShell) is not installed. Exiting"
	Return
}

# Restart Search service. Locked search indexes cause issues.
restart-service wsearch |out-null

# Get users from text file or AD. Will use text file if specified.
if ($userlist) {
	# Test if userlist file exists
	if (-NOT (test-path $userlist)) {
		write-host "File $userlist is missing"
		add-content -Path $logfile "File $userlist is missing. Exiting."
		Return
	}
	# Get Userlist from file
	$users=get-content $userlist
	# Process each user in the userlist
	foreach ($user in $users) {
		process-username $user
	}
}
else {
	# Get Userlist from specified OU in AD. Get all users if no searchbase.
	if ($searchbase) {
		$users=get-aduser -filter * -searchbase $searchbase
	} else {
		$users=get-aduser -filter *
	}

	# Process each user in the userlist
	foreach ($user in $users) {
		process-username $user.samaccountname
	}
}
#endregion Execute
