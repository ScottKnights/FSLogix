<#
    .SYNOPSIS
	Convert VHD files to VHDX files using FSLogix FRX.EXE tool.

    .DESCRIPTION
	Convert FSLogix VHD files to VHDX files.
	Gets a list of VHD files from the supplied path. Will attempt to migrate them all unless selectvhd switch is specified.
	Will copy the ACL (permissions and owner) from the VHD file to the VHDX file.
	Run on a server with FSLogix installed.

    .PARAMETER profpath
        Mandatory. Target path that hosts the FSLogix Profile Containers. Alias -PP.

    .PARAMETER logfile
	Target path to logfile. Defaults to .\profconv.log. Alias -LF.

    .PARAMETER encoding
	Encoding of the log file. Default is UNICODE.

    .PARAMETER frxpath
	Path for FRX.EXE. Default is C:\Program Files\FSLogix\Apps\frx.exe. Should not need to change this. Alias -FP.

    .PARAMETER selectvhd
	Switch to allow interactive selection of profiles to process from the user list. Alias -SV.

    .PARAMETER deletevhd
	Switch to delete VHD file after it is successfully processed. Alias DV.

    .PARAMETER replacevhdx
	Switch to delete and replace the VHDX file if it already exists. Alias RV.

    .PARAMETER verbosefrx
	Switch to display verbose output from FRX.EXE. Alias -VF.

    .PARAMETER noprogress
	Switch to disable progress bar. Alias -NP

    .INPUTS
	None

    .OUTPUTS
	Logfile written to specified location. Default location is .\profconv.log

    .NOTES
	Version:	1.20210624.1
	Author:		Scott Knights

    .EXAMPLE
	.\convert-vhdtovhdx.ps1 -profpath "\\server\share"

	Description:
	Attempt to convert all VHD files in \\server\share to VHDX files.
	Do not create a VHDX if one already exists.
	Do not delete the VHD file after conversion.

    .EXAMPLE
	.\convert-vhdtovhdx.ps1 -profpath "\\server\share" -selectvhd -replacevhdx -deletevhd -verbosefrx -noprogress

	Description:
	Select from list of VHD files in \\server\share with GUI selector and attempt to convert them to VHDX files.
	Replace the VHDX if one already exists.
	Delete the VHD file after successful conversion.
	Verbose output from FRX.EXE.
	Do not show the progress bar.

#>

# ============================================================================
#region Parameters
# ============================================================================
Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias("PP")]
    [String] $profpath,

    [Parameter()]
    [Alias("FP")]
    [String] $frxpath="c:\Program Files\FSLogix\Apps\frx.exe",

    [Parameter()]
    [Alias("log","LF")]
    [String] $logfile=".\profconv.log",

    [Parameter()]
    [Alias("SV")]
    [switch] $selectvhd,

    [Parameter()]
    [Alias("RV")]
    [switch] $replacevhdx,

    [Parameter()]
    [Alias("DV")]
    [switch] $deletevhd,

    [Parameter()]
    [Alias("VF")]
    [switch] $verbosefrx,

    [Parameter()]
    [Alias("NP")]
    [switch] $noprogress,

    [Parameter()]
    [String] $encoding="unicode"
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
			$iobject|add-content $logfile -passthru -Encoding $encoding |write-output
	        }
	}
}

function convert-profile {
	param (
		[Parameter(Mandatory=$true)]
		$container
	)

	[string]$vhd=$container.fullname
	[string]$vhdx=$vhd.tolower()+"x"
	"Processing $vhd."|out-log

	# Get permissions and owner from VHD file
	$acl=get-acl $vhd

	# Check if VHDX already exists. Return if it does and replacevhdx is not selected.
	if (test-path $vhdx) {
		if ($replacevhdx) {
			"VHDX file already exists. Replace VHDX selected. Deleting file."|out-log
			remove-item $vhdx -erroraction silentlycontinue |out-null
		} else {
			"VHDX file already exists. Replace VHDX not selected. Returning."|out-log
			return
		}
	}


	# Return if VHD file is locked
	try {
		[IO.File]::OpenWrite($VHD).close()
	} catch {
		"VHD file is locked. Returning."|out-log
		return
	}

	# Convert profile using FRX.EXE
	"Migrating VHD file to VHDX."|out-log
	try {
		if ($verbosefrx) {
			& $frxpath migrate-vhd -src $vhd -Dest $vhdx -verbose|out-log
		} else {
			& $frxpath migrate-vhd -src $vhd -Dest $vhdx 2>&1 |out-null
		}
		$returncode=$lastexitcode
		if ($returncode -ne 0) {
			"FRX return code did not equal 0. Migration may have failed."|out-log
			return
		} else {
			if ($deletevhd) {
				"Migration successful. Delete VHD selected. Deleting VHD file."|out-log
				remove-item $vhd -erroraction silentlycontinue |out-null
			}
		}
	} catch {
		"$username - FRX returned an error - Investigate."|out-log
		return
	}

	# Set permissions and owner on VHDX file.
	set-acl -path $vhdx -aclobject $acl
}

#endregion Functions

# ============================================================================
#region Execute
# ============================================================================
clear-host
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
"selectvhd:	$selectvhd",
"deletevhd:	$deletevhd",
"replacevhdx:	$replacevhdx",
"verbosefrx:	$verbosefrx",
"frxpath:	$frxpath",
""
)
$params|out-log

# Test if FRX.EXE exists
if (-NOT (test-path $frxpath)) {
	"FRX.EXE is missing. FSLogix is probably not installed. Exiting."|out-log
	Return
}

# Test if container share path exists
if (-NOT (test-path $profpath)) {
	"FSLogix container share path is invalid. Exiting."|out-log
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

# Get VHDs from profile path
$tempcontainers=get-childitem $profpath -recurse | where-object {$_.extension -eq ".vhd"}

# Allow interactive selection of VHD files to convert if -selectvhd switch used
If ($selectvhd) {
	$containers=$tempcontainers|out-gridview -OutputMode Multiple -title "Select VHD files(s) to convert"
} else {
	$containers=$tempcontainers
}

# Initialise progress bar
[nullable[double]]$secondsRemaining = $null
$start=get-date
$secondsElapsed = (Get-Date) - $start
$numvhd=$containers.count
$counter=0

if ($noprogress) {
	$global:ProgressPreference="silentlycontinue"
} else {
	$global:ProgressPreference="continue"
}

foreach ($container in $containers) {

	# Progress bar
	$counter++
	$percentComplete=($counter / $numvhd) * 100
	$progressParameters = @{
        	Activity = "Progress: $counter of $numvhd $($secondsElapsed.ToString('hh\:mm\:ss'))"
	        Status = "Converting VHD files"
	        CurrentOperation = "Converting file "+$container.fullpath
	        PercentComplete = $percentComplete
	}
	if ($secondsRemaining) {
        	$progressParameters.SecondsRemaining = $secondsRemaining
	}

	Write-Progress @progressParameters

	convert-profile $container
	" "|out-log

	# Guestimate the time remaining
	$secondsElapsed = (Get-Date)-$start
	$secondsRemaining = ($secondsElapsed.TotalSeconds / $counter) * ($numvhd - $counter)
}

$global:ProgressPreference=$InitProgressPreference

#endregion Execute
