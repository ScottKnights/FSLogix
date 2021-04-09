<#
    .SYNOPSIS
	Convert user profiles to FSLogix Containers.

    .DESCRIPTION
	Converts profiles from a server to FSLogix containers using FRX.EXE.
	Will extract a list of usernames from the ProfileList registry key, unless the -userlist or -searchbase options are specified.
	The -userlist option specifies a text file containing a list of samAccountnames to convert.
	The -searchbase option species an OU containing the list of user accounts to convert. Set to root to convert all user accounts in the domain.
	If both the -userlist and -searchbase options are specified, -userlist will be used and -searchbase ignored.
	Once a user's profile has been converted, the user will be made the owner and granted full control of the folder and VHD/VHDX file.
	Run on the terminal server containing the profiles.
	Only tested with local profiles but should work with other profile solutions as long as a copy is left on the server after logoff.
	Will not attempt to convert the user's profile if any of the following are true:
		User does not exist in AD
		User is logged on
		User does not have a profile on the server
		NTUSER.DAT is locked
		User already has a VHD file

	Active Directory module for Windows (RSAT-AD-PowerShell) must be installed.
	Script will exit if it is not unless the -installrsat switch is specified, then it will attempt to install it.

	If an AD group is specified, the user will be added to the group after successful conversion. Allows for selective enablement of FSLogix for converted profiles.

	Occasionally FRX will return a non zero error code for no discernible reason and fail the conversion.
	Due to this, the script will try the conversion twice as it often succeeds on the 2nd attempt.
	If FRX still returns non zero after the second attempt, try running the script with -verbosefrx for verbose output from FRX.EXE.

    .PARAMETER vhdpath
        Mandatory. Target path that hosts the FSLogix Profile Containers. Alias -VP.

    .PARAMETER logfile
	Target path to logfile. Defaults to .\profmig.log. Alias -LF.

    .PARAMETER userlist
	Target path to text file containing list of samAccountname's to process. Alias -UL.

    .PARAMETER searchbase
	DN of OU containing the users to process. Ignored if userlist is specified. Set to "root" to process all user accounts in the domain. Alias -SB.

    .PARAMETER adgroup
	Name of active directory group to add users to after their profile is successfully converted. Alias -AG.

    .PARAMETER vhdsize
	Maximum size of the VHD/VHDX file in MB. Default is 300000. Alias -VS.

    .PARAMETER vhdx
	Switch to create VHDX files instead of VHD. Alias -VX.

    .PARAMETER frxpath
	Path for FRX.EXE. Default is C:\Program Files\FSLogix\Apps\frx.exe. Should not need to change this. Alias -FP.

    .PARAMETER cleanorphan
	Switch to clean orphaned profiles (delete registry profile keys with no matching profile folder). Alias -CO.

    .PARAMETER reversefolder
	Switch to create container folder name in reversed <USERNAME><SID> format. Alias -RF.

    .PARAMETER verbosefrx
	Switch to display verbose output from FRX.EXE. Alias -VF.

    .PARAMETER installrsat
	Switch to try to install the RSAT-AD-PowerShell feature if it is missing. Alias -IR.

    .PARAMETER noprogress
	Switch to disable progress bar. Alias -NP

    .PARAMETER encoding
	Encoding of the log file. Default is UNICODE.

    .INPUTS
	None

    .OUTPUTS
	Logfile written to specified location. Default location is .\profmig.log

    .NOTES
	Version:	2.20210406.1
	Author:		Scott Knights
	Changes from V1:
		Configure using parameters instead of in script variables.
		Improved error checking.
		Added progress bar, because progress bars are cool. Can be suppressed with -noprogress.
		Fixed some naughties spotted by PSScriptAnalyzer (unapproved verb, using aliases, etc).
		Added default option to get users from profiles on the server.
			This was how I originally wanted to get the usernames.
			Couldn't work it out before so used AD instead, then added the text file option for testing.
			Since V1 had a head slap moment realising I can just pull the SIDs from the ProfileList key.
		Added option to clean orphaned profile keys. Not really FSLogix related but we are looking there anyway.
		Added option to install AD RSAT feature.
		Added option to select VHD or VHDX.
		Added option to select forward or reversed folder naming.
		Added option for verbose FRX output.

	Version:	2.20210407.1
		Function to output to screen and log instead of using Tee-Object.
			Cannot select encoding with Tee-Object and it varies between Powershell versions, producing inconsistent logfile.
		Added encoding parameter

    .EXAMPLE
        .\convert-profiles -vhdpath \\server\share

	Description:
	Processes users with a profile on the server the script is running from.
	If the user has a profile on the server, it is converted to a VHD file in \\server\share.
	VHD folder is created in <SID>.<USERNAME> format.


    .EXAMPLE
        .\convert-profiles -vhdpath \\server\share -searchbase "root"

	Description:
	Processes all users from the active directory domain.
	If the user has a profile on the server, it is converted to a VHD file in \\server\share.
	VHD folder is created in <SID>.<USERNAME> format.

    .EXAMPLE
        .\convert-profiles -vhdpath \\server\share -searchbase "OU=Users,OU=Company,DC=domain,DC=com" -vhdx -reversefolder -adgroup "RL-FSLogix Users"

	Description:
	Processes users from organisational unit OU=Users,OU=Company,DC=domain,DC=com.
	If the user has a profile on the server, it is converted to a VHDX file in \\server\share.
	VHDX file is created in reversed <USERNAME>.<SID> format.
	Users are added to AD group RL-FSLogix Users after they are successfully processed.

    .EXAMPLE
        .\convert-profiles -vhdpath \\server\share -userlist ".\migusers.txt" -verbosefrx -installrsat -cleanorphan

	Description:
	Processes users from text file ".\migusers.txt".
	If the user has a profile on the server, it is converted to a VHD file in \\server\share
	Verbose output from FRX.EXE is shown.
	If the Windows feature RSAT-AD-PowerShell is missing, try to install it.
	Orphaned profile registry keys (no matching profile folder) are deleted.
#>

# ============================================================================
#region Parameters
# ============================================================================
Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias("VP")]
    [String] $vhdpath,

    [Parameter()]
    [Alias("UL")]
    [String] $userlist,

    [Parameter()]
    [Alias("log","LF")]
    [String] $logfile=".\profmig.log",

    [Parameter()]
    [Alias("SB")]
    [String] $searchbase,

    [Parameter()]
    [Alias("AG")]
    [String] $adgroup,

    [Parameter()]
    [Alias("VS")]
    [int32] $vhdsize=300000,

    [Parameter()]
    [Alias("VX")]
    [switch] $vhdx,

    [Parameter()]
    [Alias("FP")]
    [String] $frxpath="c:\Program Files\FSLogix\Apps\frx.exe",

    [Parameter()]
    [Alias("CO")]
    [switch] $cleanorphan,

    [Parameter()]
    [Alias("RF")]
    [switch] $reversefolder,

    [Parameter()]
    [Alias("VF")]
    [switch] $verbosefrx,

    [Parameter()]
    [Alias("IR")]
    [switch] $installrsat,

    [Parameter()]
    [Alias("NP")]
    [switch] $noprogress,

    [Parameter()]
    [String] $encoding="unicode"
)
#endregion Parameters

# ============================================================================
#region Functions
# ============================================================================
# Write inputobject to screen and log file. Allows specifying encoding, unlike Tee-Object.
function out-log {
	param ( [Parameter(ValueFromPipeline = $true)]
		[PSCustomObject[]]
		$iobjects
	)

	process {
		foreach ($iobject in $iobjects)
	        {
			$iobject|add-content $logfile -passthru -Encoding $encoding |write-output
	        }
	}
}

function convert-profile {

	# Get Username as mandatory parameter
	param (
		[Parameter(Mandatory=$true)]
		[string]$username
	)

	# Get SID of user. Return if user does not exist.
	try {
		$SID=(get-aduser -identity $username).sid.value
	} catch {
		"$username - Not found in Active Directory."|out-log
		Return
	}

	# Return if user is logged on - Anyone know a better way to test for this? This works, but make me feel a bit dirty.
	$ErrorActionPreference="Stop"
	try {
		query user $username 2>&1|out-null
		"$username - Currently logged in."|out-log
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
		"$username - No profile path value in the registry."|out-log
		return
	}

	# Return if the profile path folder does not exist. Delete orphaned key if $cleanorphan is true
	$profpath=$profsrc.ProfileImagePath
	if (-NOT (test-path $profPath)) {
		"$username - Profile path folder does not exist. Orphaned Profile."|out-log
		if ($cleanorphan) {
			remove-item -path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -force -erroraction silentlycontinue|out-null
		}
		Return
	}

	# Return if NTUSER.DAT is locked
	$ntuserdat="$profPath\ntuser.dat"
	try {
		[IO.File]::OpenWrite($ntuserdat).close()
	 } catch {
		"$username - NTUSER.DAT is locked."|out-log
		return
	}

	# Form folder and VHD names
	if ($reversefolder) {
		$fldname=$vhdpath+"\"+$username+"_"+$SID
	} else {
		$fldname=$vhdpath+"\"+$SID+"_"+$username
	}

	if ($vhdx) {
		$vhdname="Profile_"+$username+".vhdx"
	} else {
		$vhdname="Profile_"+$username+".vhd"
	}
	$fullpath=$fldname+"\"+$vhdname

	# Return if container file already exists
	if (test-path $fullpath) {
		"$username - FSLogix container already exists."|out-log
		Return
	}

	"$username - Converting profile to container."|out-log
	# Create Folder
	mkdir $fldname -erroraction silentlycontinue |out-null

	# Set Folder Owner
	$usracct=New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $env:userdomain"\"$username
	$fldacl=get-acl $fldname
	$fldacl.setowner($usracct)
	set-acl -path $fldname -aclobject $fldacl

	# Convert profile using FRX.EXE
	try {
		if ($verbosefrx) {
			& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" -verbose|out-log
		} else {
			& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" 2>&1 |out-null
		}
		$returncode=$lastexitcode
		if ($returncode -ne 0) {
			"$username - FRX return code did not equal 0. Trying again."|out-log
			& $frxpath copy-profile -filename "$fullpath" -sid "$sid" -dynamic 1 -size-mbs="$vhdsize" 2>&1 |out-null
			$returncode=$lastexitcode
			if ($returncode -ne 0) {
				"$username - FRX return code still did not equal 0 - Investigate."|out-log
				remove-item $fldname -erroraction silentlycontinue |out-null
				return
			}
		}
	} catch {
		"$username - FRX returned an error - Investigate."|out-log
		remove-item $fldname -erroraction silentlycontinue |out-null
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
			Add-ADGroupMember -identity $adgroup -members $username -erroraction stop
		} catch {
			"$username - cannot add user to group $adgroup."|out-log
		}
	}
}
#endregion Functions

# ============================================================================
#region Execute
# ============================================================================
clear-host
$ProgressPreference="silentlycontinue"

# Check that we can create the log file
$start=get-date
try {
	set-Content -Path $logfile -encoding $encoding -Value "Start run at $start" -erroraction stop
} catch {
	write-output "Cannot create log file $logfile. Check path and permissions. Exiting."
	Return
}

$params=@(
"",
"Parameters",
"logfile:	$logfile",
"Encoding:	$encoding",
"vhdpath:	$vhdpath",
"userlist:	$userlist",
"searchbase:	$searchbase",
"adgroup:	$adgroup",
"vhdsize:	$vhdsize",
"vhdx:		$vhdx",
"cleanorphan:	$cleanorphan",
"reversefolder:	$reversefolder",
"verbosefrx:	$verbosefrx",
"installrsat:	$installrsat",
"noprogress:	$noprogress",
"frxpath:	$frxpath",
""
)
$params|out-log

# Test if container share path exists
if (-NOT (test-path $vhdpath)) {
	"FSLogix container share path is invalid. Exiting."|out-log
	Return
}

# Check that we can write to the container share path
try {
	new-item -path $vhdpath -name "testfile" -force -erroraction stop|out-null
	remove-item -path $vhdpath"\testfile" -force -erroraction stop|out-null
} catch {
	"Cannot write to the FSLogix container share. Check path and permissions. Exiting."|out-log
	Return
}

# Test if FRX.EXE exists
if (-NOT (test-path $frxpath)) {
	"FRX.EXE is missing. FSLogix is probably not installed. Exiting."|out-log
	Return
}

# Test if RSAT-AD-PowerShell is installed. If not and the $installrsat switch is true, try to install it.
$rsatad=Get-WindowsFeature *RSAT-AD-PowerShell*
if (-not ($rsatad.installed)) {
	if ($installrsat) {
		"RSAT-AD-PowerShell not installed. Install requested. Attempting installation."|out-log
		try {
			$ProgressPreference="continue"
			Install-WindowsFeature RSAT-AD-PowerShell|out-null
			$ProgressPreference="silentlycontinue"
		} catch {
			"Error installing RSAT-AD-PowerShell."|out-log
		}
	}
}

# Test if RSAT-AD-PowerShell is installed (again!).
$rsatad=Get-WindowsFeature *RSAT-AD-PowerShell*
if (-not ($rsatad.installed)) {
	"Active Directory module for Windows (RSAT-AD-PowerShell) is not installed. Exiting."|out-log
	Return
} else {
	import-module activedirectory
	$ProgressPreference="continue"
}

# Restart Search service. Locked search indexes cause issues.
restart-service wsearch -erroraction silentlycontinue|out-null

# Get users from local profiles, text file or AD.
if ($userlist) {
	# Get list of usernames from a text file. Useful for testing/piloting.
	# Test if userlist file exists
	"Getting userlist from text file $userlist."|out-log
	if (-NOT (test-path $userlist)) {
		"File $userlist is missing. Exiting."|out-log
		Return
	}
	# Get Userlist from file
	$users=get-content $userlist
} elseif ($searchbase) {
	# Get Userlist from specified OU in AD. Get all users if searchbase = root.
	# Probably not so useful now we have parselocal, but already in place so will leave it.
	"Getting userlist from Active Directory, Searchbase $searchbase."|out-log
	try {
		if ($searchbase -eq "root") {
			$userobjects=get-aduser -filter * -erroraction stop
		} else {
			$userobjects=get-aduser -filter * -searchbase $searchbase -erroraction stop
		}
	} catch {
		"Unable to get users from Active Directory. Exiting."|out-log
	}
	$users=$userobjects.samaccountname
} else {
	# Get Userlist by parsing SIDs from ProfileList key. Probably most useful option.
	# Yes, we are converting SID to username, then converting it back to a SID again in the function.
	# Thought of just passing the SID, but that would have meant too much messing with the function.
	"Getting userlist from ProfileList key in the registry."|out-log
	$domain=$env:userdomain
	$users=@()
	$keys=(Get-ChildItem -path "hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList").name
	foreach ($key in $keys) {
		$sid=$key.Replace("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\", [string]::Empty)
		$objSID = New-Object System.Security.Principal.SecurityIdentifier ` ($sid)
		try {
			$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
		} catch {
			"SID $SID is invalid."|out-log
		}
		# We only want domain accounts
		$username=$objUser.Value
		if ($username -match $domain) {
			$user=$username.replace("$domain\",[string]::Empty)
			$users += $user
		}
	}
}

# Initialise progress bar
[nullable[double]]$secondsRemaining = $null
$start=get-date
$secondsElapsed = (Get-Date) - $start
$numusers=$users.count
$counter=0

# Process each user in the userlist
foreach ($user in $users) {
	# Progress bar
	$counter++
	$percentComplete=($counter / $numusers) * 100
	$progressParameters = @{
        	Activity = "Progress: $counter of $numusers $($secondsElapsed.ToString('hh\:mm\:ss'))"
	        Status = "Converting Profiles"
	        CurrentOperation = "Converting profile for user "+$User
	        PercentComplete = $percentComplete
	}
	if ($secondsRemaining) {
        	$progressParameters.SecondsRemaining = $secondsRemaining
	}
	if (-NOT ($noprogress)) {
		Write-Progress @progressParameters
	}

	convert-profile $user

	# Guestimate the time remaining
	$secondsElapsed = (Get-Date)-$start
	$secondsRemaining = ($secondsElapsed.TotalSeconds / $counter) * ($numusers - $counter)
}
#endregion Execute