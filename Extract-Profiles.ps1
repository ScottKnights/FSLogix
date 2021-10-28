<#
    .SYNOPSIS
	Extract data from VHD/VHDX files

    .DESCRIPTION
	Extract selected folders from FSLogix profile containers.

    .PARAMETER profpath
        Mandatory. Target path that hosts the FSLogix Profile Containers. Alias -PP.

    .PARAMETER destpath
        Mandatory. Path for the extracted profiles. Alias -DP.

    .PARAMETER driveletter
	Drive letter to map the container to. Defaults to P. Select something else if P is already in use

    .PARAMETER folders
        Array of folders to extract. Default is documents and downloads.

    .PARAMETER selectprofiles
	Switch to allow interactive selection of profiles to process from the user list.

    .PARAMETER logfile
	Target path to logfile. Defaults to .\exprof.log. Alias -LF.

    .INPUTS
	None

    .OUTPUTS
	Logfile written to specified location. Default location is .\exprof.log

    .NOTES
	Version:	1.20211028.1
	Author:		Scott Knights

    .EXAMPLE
	.\extract-profile.ps1 -profpath "\\server\share" -destpath "d:\extract"

	Description:
	Extract files from all VHD/VHDX files in \\server\share to d:\extract.

#>

# ============================================================================
#region Parameters
# ============================================================================
Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias("PP")]
    [String] $profpath,

    [Parameter(Mandatory=$true,Position=1)]
    [Alias("DP")]
    [String] $destpath,

    [Parameter()]
    [Alias("SP")]
    [switch] $selectprofiles,

    [Parameter()]
    [String[]] $folders=@("Documents","Downloads"),

    [Parameter()]
    [Alias("DL")]
    [String] $driveletter="P",

    [Parameter()]
    [Alias("log","LF")]
    [String] $logfile=".\exprof.log"

)
#endregion Parameters

# ============================================================================
#region Functions
# ==========================================================================
function out-log {
	param ( [Parameter(ValueFromPipeline = $true)]
		[PSCustomObject[]]
		$iobjects
	)

	process {
		foreach ($iobject in $iobjects)
	        {
			$iobject|add-content $logfile -Encoding $encoding
	        }
	}
}

function extract-profile {
	param (
		[Parameter(Mandatory=$true)]
		$container
	)

	[string]$vhd=$container.fullname
	[string]$profile=$container.name -replace "Profile_", "" -replace ".vhdx","" -replace ".vhd",""
	# Set diskpart commands
	$attach = "sel vdisk file=`"$vhd`"`r`nattach vdisk"
	$assign = "sel vdisk file=`"$vhd`"`r`nsel part 1`r`nassign letter=$driveletter"
	$detach = "sel vdisk file`"$vhd`"`r`ndetach vdisk"
	"Processing $vhd."|out-log

	# Extract profile 

	# Mount the VHD and assign a drive letter
	$attach | diskpart |out-log
	Start-Sleep -s 2
	$assign | diskpart |out-log

	# Create the destination path and copy files

	Foreach ($folder in $folders) {
		$srcpath="$driveletter"+":\profile\"+$folder
		robocopy $srcpath "$destpath\$profile\$folder\" /e /copy:dat /xj /xf "desktop.ini" "thumbs.db" /r:1 /w:1 /np |out-log
	}

	# Detach the VHD
	$detach | diskpart |out-log

}

#endregion Functions

# ============================================================================
#region Execute
# ============================================================================
clear-host
$encoding="ASCII"
$InitProgressPreference=$global:ProgressPreference
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
"profpath:	$profpath",
"destpath:	$destpath",
"selectprofiles:	$selectprofiles",
"folders:	$folders",
"driveletter:	$driveletter",
""
)
$params|out-log

# Test if container share path exists
if (-NOT (test-path $profpath)) {
	"FSLogix container share path is invalid. Exiting."|out-log
	Return
}

# Test if destination path exists
if (-NOT (test-path $destpath)) {
	"Destination path is invalid. Exiting."|out-log
	Return
}


# Check that we can write to the container share path
try {
	new-item -path $profpath -name "testfile" -force -erroraction stop|out-null
	remove-item -path $profpath"\testfile" -force -erroraction stop|out-null
} catch {
	"Cannot write to the FSLogix container share. Check path and permissions. Exiting."|out-log
	Return
}

# Check that we can write to the destination path
try {
	new-item -path $destpath -name "testfile" -force -erroraction stop|out-null
	remove-item -path $destpath"\testfile" -force -erroraction stop|out-null
} catch {
	"Cannot write to the destination path. Check path and permissions. Exiting."|out-log
	Return
}

# Get VHD/VHDXs from profile path
$tempcontainers=get-childitem $profpath -recurse | where-object {($_.Extension -Match ".vhd" -or $_.Extension -eq ".vhdx")}

# Allow interactive selection of VHD files to convert if -selectvhd switch used
If ($selectprofiles) {
	$containers=$tempcontainers|out-gridview -OutputMode Multiple -title "Select containers to extract from"
} else {
	$containers=$tempcontainers
}

# Initialise progress bar
[nullable[double]]$secondsRemaining = $null
$start=get-date
$secondsElapsed = (Get-Date) - $start
$numvhd=$containers.count
$counter=0

$global:ProgressPreference="continue"

foreach ($container in $containers) {

	# Progress bar
	$counter++
	$percentComplete=($counter / $numvhd) * 100
	$progressParameters = @{
        	Activity = "Progress: $counter of $numvhd $($secondsElapsed.ToString('hh\:mm\:ss'))"
	        Status = "Extracting Profile Data"
	        CurrentOperation = "Extracting from "+$container.fullname
	        PercentComplete = $percentComplete
	}
	if ($secondsRemaining) {
        	$progressParameters.SecondsRemaining = $secondsRemaining
	}

	Write-Progress @progressParameters

	extract-profile $container
	" "|out-log

	# Guestimate the time remaining
	$secondsElapsed = (Get-Date)-$start
	$secondsRemaining = ($secondsElapsed.TotalSeconds / $counter) * ($numvhd - $counter)
}

$global:ProgressPreference=$InitProgressPreference

#endregion Execute