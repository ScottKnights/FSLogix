<#
    .SYNOPSIS
	Export selected files/folders from FSLogix VHD/VHDX files.

    .DESCRIPTION
	Export selected files/folders from FSLogix profile containers.

	Mount each profile container discovered in the source profile path as a drive letter using Diskpart.
	If the SelectProfiles switch is specified, the discovered containers are displayed in GridView. Only selected profiles will be processed.
	After each container is mounted, the selected files/folders will be copied to a subfolder in the destination path using RoboCopy.

    .PARAMETER ProfilePath
        Mandatory. Target path that hosts the FSLogix Profile Containers.

    .PARAMETER DestinationPath
        Mandatory. Path for the exported profile data.

    .PARAMETER DriveLetter
	Drive letter to map the container to. Defaults to P. Select something else if P is already in use.

    .PARAMETER Recurse
	Switch to recurse folders in robocopy. Adds /S option to the robocopy command. Takes precedence over RecurseEmpty if both are specified.

    .PARAMETER RecurseEmpty
	Switch to recurse all folders in robocopy, including empty ones. Adds /E option to the robocopy command.

    .PARAMETER Folders
        Array of folder names to export. Default is documents and downloads.

    .PARAMETER Files
        Array of file names to export. Default is all files in the selected folders.

    .PARAMETER ExcludeFiles
        Array of file names to exclude from the export. Default is desktop.ini and thumbs.db.

    .PARAMETER SelectProfiles
	Switch to allow interactive selection of profiles to process from the user list.

    .PARAMETER LogFile
	Target path to logfile. Defaults to .\ExportProfiles.log.

    .INPUTS
	None

    .OUTPUTS
	Logfile written to specified location. Default location is .\ExportProfiles.log.

    .NOTES
	Scott Knights
	V 1.20211028.1
		Initial Release. Only exported folders. Always recursed.
	V 2.20230522.1
		Added option to also select files to include/exclude and whether to recurse.
		Rename to use approved verb.
	V 2.20230523.1
		Comment changes and tidy up.

    .EXAMPLE
	Export-DataFromProfiles.ps1 -ProfilePath "\\server\share" -DestinationPath "d:\export" -RecurseEmpty -DriveLetter K -Logfile C:\Temp\Mylog.txt

	Description:
	Export files and folders (including empty ones) from the Documents and Downloads folders in all VHD/VHDX files in \\server\share to d:\export.
	Exclude Thumbs.db and Desktop.ini.
	Mount the VHD/VHDX as drive letter K
	Write the log to C:\Temp\Mylog.txt

    .EXAMPLE
	Export-DataFromProfiles.ps1 -ProfilePath D:\FSLogix -DestinationPath D:\Export -SelectProfiles -Folders 'AppData\Local\Google\Chrome\User Data\Default' -Files bookmarks,favicons

	Description:
	Select which profiles to process in source folder D:\FSlogix
	Export only the bookmarks and favicons files from folder AppData\Local\Google\Chrome\User Data\Default to D:\Export
	Do not recurse
	Mount the VHD/VHDX as default drive letter P
	Write the log to the default .\ExportProfiles.log

#>

# ============================================================================
#region Parameters
# ============================================================================
Param(
    [Parameter(Mandatory=$true,Position=0)]
    [String] $ProfilePath,

    [Parameter(Mandatory=$true,Position=1)]
    [String] $DestinationPath,

    [Parameter()]
    [switch] $SelectProfiles,

    [Parameter()]
    [switch] $Recurse,

    [Parameter()]
    [switch] $RecurseEmpty,

    [Parameter()]
    [String[]] $Folders=@("Documents","Downloads"),

    [Parameter()]
    [String[]] $Files=@(),

    [Parameter()]
    [String[]] $ExcludeFiles=@("desktop.ini","thumbs.db"),

    [Parameter()]
    [String] $DriveLetter="P",

    [Parameter()]
    [String] $LogFile=".\ExportProfiles.log"

)
#endregion Parameters

# ============================================================================
#region Functions
# ==========================================================================
Function Out-Log {
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

Function Export-Profile {
	param (
		[Parameter(Mandatory=$true)]
		$Container
	)

	[String]$Recursion=$null
	if ($Recurse) {
		$Recursion="/s"
	} elseif ($RecurseEmpty) {
		$Recursion="/e"
	}

	[String]$VHD=$Container.FullName
	[String]$Profile=$Container.Name -replace "Profile_", "" -replace ".vhdx","" -replace ".vhd",""
	# Set diskpart commands
	[String]$Attach = "sel vdisk file=`"$VHD`"`r`nattach vdisk"
	[String]$Assign = "sel vdisk file=`"$VHD`"`r`nsel part 1`r`nassign letter=$driveletter"
	[String]$Detach = "sel vdisk file`"$VHD`"`r`ndetach vdisk"
	"Processing $VHD."|Out-Log

	# Export data from profile

	# Mount the VHD and assign a drive letter
	$Attach | diskpart |Out-Log
	Start-Sleep -s 2
	$Assign | diskpart |Out-Log

	# Create the destination path and copy files

	Foreach ($Folder in $Folders) {
		[String]$SourcePath="$driveletter"+":\profile\"+$folder
		$SourcePath |Out-Log
		[String]$Destination=$DestinationPath+"\"+$Profile+"\"+$Folder
		$Destination |Out-Log
		robocopy "$SourcePath" "$DestinationPath\$Profile\$Folder" $Files $Recursion /copy:dat /xj /xf $ExcludeFiles /r:1 /w:1 /np |Out-Log
	}

	# Detach the VHD
	$Detach | diskpart |Out-Log

}

#endregion Functions

# ============================================================================
#region Execute
# ============================================================================
Clear-Host
[String]$Encoding="ASCII"
[System.Enum]$InitProgressPreference=$global:ProgressPreference
# Check that we can create the log file
[DateTime]$Start=Get-Date
Try {
	Set-Content -Path $LogFile -Encoding $Encoding -Value "Start run at $start" -ErrorAction Stop
} Catch {
	Write-Output "Cannot create log file $logfile. Check path and permissions. Exiting."
	Return
}

$Params=@(
"",
"Parameters",
"LogFile:		$LogFile",
"Encoding:		$Encoding",
"ProfilePath:		$ProfilePath",
"DestinationPath:	$DestinationPath",
"SelectProfiles:	$SelectProfiles",
"Folders:		$Folders",
"Files:			$Files",
"Recurse:		$Recurse",
"RecurseEmpty:		$RecurseEmpty",
"ExcludeFiles:		$ExcludeFiles",
"DriveLetter:		$DriveLetter",
""
)
$Params|Out-Log

# Test if container share path exists
If (-NOT (Test-Path -LiteralPath $ProfilePath)) {
	"FSLogix container share path is invalid. Exiting."|Out-Log
	Return
}

# Test if destination path exists
If (-NOT (Test-Path -LiteralPath $DestinationPath)) {
	"Destination path is invalid. Exiting."|Out-Log
	Return
}


# Check that we can write to the container share path
Try {
	New-Item -Path $ProfilePath -Name "testfile" -Force -ErrorAction Stop | Out-Null
	Remove-Item -LiteralPath $ProfilePath"\testfile" -Force -ErrorAction Stop | Out-Null
} Catch {
	"Cannot write to the FSLogix container share. Check path and permissions. Exiting." | Out-Log
	Return
}

# Check that we can write to the destination path
Try {
	New-Item -Path $DestinationPath -Name "testfile" -Force -ErrorAction Stop | Out-Null
	Remove-Item -LiteralPath $DestinationPath"\testfile" -Force -ErrorAction Stop | out-null
} Catch {
	"Cannot write to the destination path. Check path and permissions. Exiting." | Out-Log
	Return
}

# Get VHD/VHDXs from profile path
[System.Array]$TempContainers=Get-ChildItem $ProfilePath -Recurse | Where-Object {($_.Extension -Match ".vhd" -or $_.Extension -eq ".vhdx")}
[System.Array]$Containers=@()

# Allow interactive selection of VHD files to convert if -SelectProfiles switch used
If ($SelectProfiles) {
	$Containers=$TempContainers|Out-GridView -OutputMode Multiple -Title "Select containers to export data from"
} else {
	$Containers=$TempContainers
}

# Initialise progress bar
[Nullable[Double]]$SecondsRemaining = $null
[DateTime]$Start=Get-Date
[TimeSpan]$SecondsElapsed = (Get-Date) - $start
[Int32]$VHDCount=$containers.count
[Int32]$Counter=0
[Int32]$PercentComplete=0
[Double]$SecondsRemaining=0

$Global:ProgressPreference="continue"

ForEach ($Container in $Containers) {

	# Progress bar
	$Counter++
	$PercentComplete=($Counter / $VHDCount) * 100
	$ProgressParameters = @{
        	Activity = "Progress: $Counter of $VHDCount $($SecondsElapsed.ToString('hh\:mm\:ss'))"
	        Status = "Exporting Profile Data"
	        CurrentOperation = "Exporting from "+$Container.Fullname
	        PercentComplete = $PercentComplete
	}
	if ($SecondsRemaining) {
        	$ProgressParameters.SecondsRemaining = $SecondsRemaining
	}

	Write-Progress @ProgressParameters

	Export-Profile $Container
	" "|Out-Log

	# Guestimate the time remaining
	$SecondsElapsed = (Get-Date)-$Start
	$SecondsRemaining = ($SecondsElapsed.TotalSeconds / $Counter) * ($VHDCount - $Counter)
}

$global:ProgressPreference=$InitProgressPreference

#endregion Execute