<#
 Configure-VM-Export.ps1 version 1.1
 Build and manipulate XML configuration data for Export-VMs.ps1
#>

## Global "Constants" (actually variables, but are unchangeable by functions)
$ScriptDir = Split-Path (Get-Variable MyInvocation).Value.MyCommand.Path
$ConfigFile = "$ScriptDir\VMExportConfig.xml"				# Modify this to change where configuration data is saved
$DefaultExportPath = "\\mtfs\drobodata$\Hyper-V_Exports"	# Modify this to change the default location of exported VMs
$DefaultLogFile = "$ScriptDir\VMExport.log"					# Modify this to change the name of the default log file. The \ is required.
$DefaultUseCluster = "No"									# Modify this to change the default cluster search behavior
$DefaultActionBeforeExport = "Save"							# Modify this to change default pre-export action, "Save" or "Shutdown"
$DefaultMoveMissingVMsToThisNode = "Yes"					# Modify this to change default handling of missing VMs
$DefaultIncludeOrExclude = "Exclude"						# Modify this to change default processing of VM list
[int]$ObjectsPerLine = 3									# Formatting control based on standard 80-char screen width
[int]$LargeObjectsPerLine = 2								# Formatting control based on standard 80-char screen width
$VerStamp = "Configure VM Export version 1.0"
$CopyrightStamp = "TOUTON Hyper-V Core 2012 R2 VM Export"
$BackColorHeader = "White"
$ForeColorHeader = "DarkBlue"
$BackColor = "Black"
$ForeColor = "Green"
$BackColorError = "Black"
$ForeColorError = "Red"
$DefaultConfigTemplate = @"
<?xml version=`"1.0`" encoding=`"UTF-8`"?>
<Root>
	<Configuration>
		<DefaultExportPath>$DefaultExportPath</DefaultExportPath>
		<LogFile>$DefaultLogFile</LogFile>
		<UseCluster>$DefaultUseCluster</UseCluster>
		<DefaultActionBeforeExport>$DefaultActionBeforeExport</DefaultActionBeforeExport>
		<DefaultMoveMissingVMsToThisNode>$DefaultMoveMissingVMsToThisNode</DefaultMoveMissingVMsToThisNode>
		<DefaultIncludeOrExclude>$DefaultIncludeOrExclude</DefaultIncludeOrExclude>
	</Configuration>
	<Jobs />
</Root>
"@

## End global "constants"

## Global variables
$global:ConfigData = ""					#  in-memory representation of configuration data in an XML structure

## End global variables

############## Functions ##############

## Takes a string and returns true if it is numeric
function Is-StringNumeric
{
	param([String]$StringToValidate)
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null	# Seems to work without this, but just in case
	if([Microsoft.VisualBasic.Information]::IsNumeric($StringToValidate))
	{ return $TRUE }
	else
	{ return $FALSE }
}

## Opens an open file dialog and returns the selected file name
##
## $Title = What to show on the title bar of the OFD
## $InitialDirectory = the directory that the OFD will start looking in
## $Filter = the file name.type to look for
## $DefaultExt = default extension (.xxx)
function Get-FileSelection
{
	param([string]$Title,[string]$InitialDirectory=".\",[string]$Filter="All Files (*.*)|*.*",[string]$DefaultExt="")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$GetFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$GetFileDialog.InitialDirectory = $InitialDirectory
	$GetFileDialog.CheckFileExists = $FALSE
	$GetFileDialog.Filter = $Filter
	$GetFileDialog.Title = $Title
	$GetFileDialog.DefaultExt = $DefaultExt
	$GetFileDialog.ShowHelp = $TRUE			# Needed because PowerShell runs MTA by default, without it, the session will hang
	$GetFileDialog.RestoreDirectory = $TRUE	# don't change the working directory
	$DialogResult = $GetFileDialog.ShowDialog()
	if ($DialogResult -eq "OK")
	{
		return $GetFileDialog.Filename
	}
	else { return "" }
}

## Essentially the same as above. There is a select folder dialog available, but it only works when PowerShell is started
## in single-threaded apartment mode. Since it's unreasonable to expect a user to know what to do in advance, the next best
## option is to use an OpenFileDialog and ignore any file name that is selected. Not the greatest interface in the world,
## but easier on the developer than a DIY UI and easier on the user than manually typing in a path.
##
## $Title = What to show on the title bar of the OFD
## $InitialDirectory = the directory that the OFD will start looking in
function Get-FolderSelection
{
	param([string]$Title,[string]$InitialDirectory=".\")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$GetFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$GetFileDialog.InitialDirectory = $InitialDirectory
	$GetFileDialog.Filename = "Filename will be ignored"
	$GetFileDialog.CheckFileExists = $FALSE
	$GetFileDialog.Title = "Select Folder"
	$GetFileDialog.ShowHelp = $TRUE					# Needed because PowerShell runs MTA by default
	$DialogResult = $GetFileDialog.ShowDialog()
	if ($DialogResult -eq "OK")
	{
		return Split-Path $GetFileDialog.Filename
	}
	else { return "" }
}

## Read the configuration file and ensure it is more or less valid format
function Read-ConfigData
{
	$global:ConfigData = New-Object XML
	if (Test-Path $ConfigFile)	# Don't blindly read it
	{
		try
		{
			$Error.Clear()
			$global:ConfigData.Load($ConfigFile)
		}
		catch
		{
			Clear-Host
			Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError $Error
			Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "`n`nThere was a problem reading your configuration file ($ConfigFile).`nError details above."
			Create-ConfigFile
			return $FALSE # This function is being called in a loop. If the caller gets a $FALSE, it will call again, so the config file will be read
		}
	}
	else
	{
		Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "`n`nThere was a problem locating the configuration file: $ConfigFile."
		Create-ConfigFile
		return $FALSE
	}
	$ConfigurationRoot = $global:ConfigData.Root.Configuration

	# The Root element is not checked for, if the file is that badly damaged then it probably didn't even load
	if ($ConfigurationRoot -eq $null)
	{
		$global:ConfigData.SelectSingleNode("Root").AppendChild($global:ConfigData.CreateElement("Configuration"))
		$ConfigurationRoot = $global:ConfigData.SelectSingleNode("Root/Configuration")
	}

	# TODO: If the script is upgraded with new options, ensure they are checked below
	# TODO: Implement better error-checking: only checks to see if the nodes are non-present or empty, not for invalid data
	# TODO: Find a way to loop these, too much almost-duplicate code

	# Check to be sure the DefaultExportPath key exists and is set
	if ($ConfigurationRoot.DefaultExportPath -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("DefaultExportPath"))
	}
	if ($ConfigurationRoot.DefaultExportPath.Length -eq 0)
	{
		$ConfigurationRoot.DefaultExportPath = $DefaultExportPath
		$SaveNeeded = $TRUE
	}
	# Check to be sure the LogFile key exists and is set
	if ($ConfigurationRoot.LogFile -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("LogFile"))
	}
	if ($ConfigurationRoot.LogFile.Length -eq 0)
	{
		$ConfigurationRoot.LogFile = $DefaultLogFile
		$SaveNeeded = $TRUE
	}
	# Check to be sure the UseCluster key exists and is set
	if ($ConfigurationRoot.UseCluster -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("UseCluster"))
	}
	if ($ConfigurationRoot.UseCluster.Length -eq 0)
	{
		$ConfigurationRoot.UseCluster = $DefaultUseCluster
		$SaveNeeded = $TRUE
	}
	# Check to be sure the DefaultActionBeforeExport key exists and is set
	if ($ConfigurationRoot.DefaultActionBeforeExport -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("DefaultActionBeforeExport"))
	}
	if ($ConfigurationRoot.DefaultActionBeforeExport.Length -eq 0)
	{
		$ConfigurationRoot.DefaultActionBeforeExport = $DefaultActionBeforeExport
		$SaveNeeded = $TRUE
	}
	# Check to be sure the DefaultMoveMissingVMsToThisNode key exists and is set
	if ($ConfigurationRoot.DefaultMoveMissingVMsToThisNode -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("DefaultMoveMissingVMsToThisNode"))
	}
	if ($ConfigurationRoot.DefaultMoveMissingVMsToThisNode.Length -eq 0)
	{
		$ConfigurationRoot.DefaultMoveMissingVMsToThisNode = $DefaultMoveMissingVMsToThisNode
		$SaveNeeded = $TRUE
	}
	# Check to be sure the DefaultIncludeOrExclude key exists and is set
	if ($ConfigurationRoot.DefaultIncludeOrExclude -eq $null)
	{
		$ConfigurationRoot.AppendChild($ConfigData.CreateElement("DefaultIncludeOrExclude"))
	}
	if ($ConfigurationRoot.DefaultIncludeOrExclude.Length -eq 0)
	{
		$ConfigurationRoot.DefaultIncludeOrExclude = $DefaultIncludeOrExclude
		$SaveNeeded = $TRUE
	}
	if ($SaveNeeded) { Save-ConfigData }
	return $TRUE
}

## Saves the in-memory configuration to disk
function Save-ConfigData
{
	param([Boolean]$SetDefaults = $FALSE)
	if ($SetDefaults)	# not currently used in this build
	{
		Create-ConfigFile
		Read-ConfigData
		return
	}
	try
	{
		$global:ConfigData.Save($ConfigFile)
	}
	catch {
		Clear-Host
		Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "$Error`nUnable to save changes to $ConfigFile due to above error.`nPlease correct the error and run the configurator again."
		Exit
	}
}

function Create-ConfigFile
{
	$CreateNew = "u"
	while ($CreateNew -notmatch "[y|n]")	# not case sensitive
	{
		# only read the first character: if it's anything other than a y or an n or a blank line, make the user try again
		$CreateNew = (Read-Host -Prompt "`nAttempt to create a new configuration file? ($ConfigFile)`nIf the file exists but is damaged, `nTHIS WILL DESTROY ANY CONFIGURED JOBS! [Enter=No]")[0]
		if ($CreateNew -eq $null) { $CreateNew = "n" }
	}
	if ($CreateNew -eq "n")	# also not case sensitive
	{
		Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "Program cannot continue without a readable configuration file.`nPlease attempt to correct the problem preventing $ConfigFile from being accessed."
		Exit
	}

	try
	{
		$Error.Clear()
		Set-Content $ConfigFile $DefaultConfigTemplate -ErrorAction Stop	# try to put defaults into the file and bomb if that fails
	}
	catch
	{
		Clear-Host
		Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError $Error
		Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "`n`nAttempted to create $ConfigFile but failed due to above errors.`nThis is an unrecoverable problem. Exiting program."
		Exit
	}
}

## Build the pretty menu header
##
## $HeaderTitle = text to appear within the header
function Show-Header
{
	param([String]$HeaderTitle)
	$HeaderTitle = $HeaderTitle
	$WindowWidth = $Host.UI.RawUI.WindowSize.Width - 1
	$LineBar = " " + ("-" * ($WindowWidth - 2)) + " `n"
	Clear-Host
	$Host.UI.RawUI.WindowTitle = $HeaderTitle
	$Header = ($LineBar +
		("|" + $HeaderTitle.PadLeft(($HeaderTitle.Length / 2) + ($WindowWidth / 2)).PadRight($WindowWidth - 2) + "|" + "`n") + 
		("|" + "".PadLeft($WindowWidth / 2).PadRight($WindowWidth - 2) + "|" + "`n") +
		("|" + $CopyrightStamp.PadLeft(($CopyrightStamp.Length / 2) + ($WindowWidth / 2)).PadRight($WindowWidth - 2) + "|`n") +
		$LineBar )
	Write-Host -ForegroundColor $ForeColorHeader -BackgroundColor $BackColorHeader $Header
}

## Main menu
## The user's response is returned from most menus, but currently nothing is really done with the return itself
function Show-MainMenu
{
	Show-Header $VerStamp
	$MenuText =
		" [1] Change global options`n" +
		" [2] Work with export job definitions`n" +
		" [3] Clear job log`n" +
		"`n [0] Exit Configuration Tool"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host -Prompt "`nChoose an option"

	switch ($MenuChoice)
	{
		0 { return 0 }
		1	{
			while (Show-GlobalOptionsMenu) {}
			return 1
		}
		2 {
			while (Show-JobOptionsMenu) {}
			return 2
		}
		3 {
			Clear-Host
			Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor "Proceeding will clear all data in the log file ($($global:ConfigData.Root.Configuration.LogFile))."
			$ClearResponse = (Read-Host -Prompt "Do you wish to proceed? [Enter=No]")[0]	# Notice the [0]: only the first character is trapped
			if ($ClearResponse -eq "y")
			{
				try
				{
					$stamp = Get-Date
					$uid =  [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
					Set-Content $global:ConfigData.Root.Configuration.LogFile "Log cleared on $stamp by $uid" -ErrorAction Stop
				}
				catch
				{
					Clear-Host
					Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError $Error
					Write-Host -ForegroundColor $ForeColorError -BackgroundColor $BackColorError "`nUnable to clear the log file due to the above error.`nThis condition may also prevent writing to the log file.`nThe program will now exit to allow for investigation of the file problem.`nIf the issue cannot be corrected, use the configurator to choose a new log file."
					Exit
				}
			}
			return 3
		}
		default { return -1 }	# Will cause the menu to recycle
	}
}

## Options on this menu are global to the application. See the "VMExportConfig.xml" file for clarity
function Show-GlobalOptionsMenu
{
	$ConfigNode = $global:ConfigData.SelectSingleNode("Root/Configuration")
	$UseCluster = $ConfigNode.UseCluster
	if ($UseCluster -eq "No")
	{	$LMOption = "Disabled" }
	else { $LMOption = $ConfigNode.DefaultMoveMissingVMsToThisNode }
	Show-Header "Configure VM Export Global Options"
	$MenuText =
		" [1]	Change Default Export Path`n" +
		"	[" + $ConfigNode.DefaultExportPath + "]`n" +
		" [2]	Change Log File Name`n" +
		"	[" + $ConfigNode.LogFile + "]`n" +
		" [3]	Save or shut down VMs prior to export?`n" +
		"	[" + $ConfigNode.DefaultActionBeforeExport + "]`n" +
		" [4]	Use cluster?`n" +
		"	[" + $ConfigNode.UseCluster + "]`n" +
		" [5]	Attempt to LiveMigrate missing VMs to this node?`n" +
		"	[$LMOption]`n" +
		" [6]	Include or exclude selected VMs by default?`n" +
		"		Include will export only selected VMs`n" +
		"		Exclude will export all VMs except those that are selected`n" +
		"	[" + $ConfigNode.DefaultIncludeOrExclude + "]`n" +
		"`n [0]	Exit To Main Menu"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host -Prompt "`nChoose an option"

	switch ($MenuChoice)
	{
		0	{
			Save-ConfigData
			return 0
		}
		1 {
			$NewPath = Get-FolderSelection -Title "Select Default Export Path" -InitialDirectory $ConfigNode.DefaultExportPath
			if ($NewPath.Length -gt 0)
			{
				if (Test-Path($NewPath))
				{
					$ConfigNode.DefaultExportPath = "$NewPath"
				}
			}
			return 1
		}
		2 {
			$NewName = Get-FileSelection -Title "Select Log Filename" -InitialDirectory $ConfigNode.LogFilePath -Filter "Log Files (*.log)|*.log" -DefaultExt "log"
			if ($NewName.Length -gt 0)
			{
				if (Test-Path(Split-Path($NewName)))
				{
					$ConfigNode.LogFile = "$NewName"
				}
			}
			return 2
		}
		3 {
			if ($ConfigNode.DefaultActionBeforeExport -eq "Save" )
			{ $ConfigNode.DefaultActionBeforeExport = "Shutdown" }
			else
			{ $ConfigNode.DefaultActionBeforeExport = "Save" }
			return 3
		}
		4 {
			if ($ConfigNode.UseCluster -eq "Yes" )
			{ $ConfigNode.UseCluster = "No" }
			else
			{ $ConfigNode.UseCluster = "Yes" }
			return 4
		}
		5 {
			if ($LMOption -ne "Disabled")
			{
				if ($ConfigNode.DefaultMoveMissingVMsToThisNode -eq "Yes")
				{ $ConfigNode.DefaultMoveMissingVMsToThisNode = "No" }
				else
				{ $ConfigNode.DefaultMoveMissingVMsToThisNode = "Yes" }

			}
			return 5
		}
		6 {
				if ($ConfigNode.DefaultIncludeOrExclude -eq "Include")
				{ $ConfigNode.DefaultIncludeOrExclude = "Exclude" }
				else
				{ $ConfigNode.DefaultIncludeOrExclude = "Include" }
				return 6
			}
		default { return -1 }	# Will cause the menu to recycle
	}
}

## Top-level menu for changing job definitions
function Show-JobOptionsMenu {
	Show-Header "Configure Job Definitions"
	$MenuText =
		" [1]	Add/modify export job definition`n" +
		" [2]	Delete export job definition`n" +	# TODO: hide this when there are no jobs to delete
		"`n [0]	Exit To Main Menu"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host -Prompt "`nChoose an option"
	switch ($MenuChoice)
	{
		0	{ return 0 }
		1 {
			while (Show-AddModifyJobMenu) {}
			return 1
		}
		2 {
			while (Show-DeleteJobMenu) {}
			return 2
		}
		default { return -1 }	# Will cause the menu to recycle
	}
}

## Menu for deletion of entire job definitions
function Show-DeleteJobMenu {
	$JobNodes = $global:ConfigData.SelectNodes("/Root/Jobs/Job")
	Show-Header "Delete Export Definitions"
	$NameList = @{}
	$MenuText = ""
	ForEach ($JobNode in $JobNodes)
	{
		$Base++
		$NameList.Add("$Base", $JobNode.JobName)
		$MenuText += " [$Base]	" + $JobNode.JobName + "`n"
	}
	if ($JobNodes.Count -eq 0)
	{ $MenuText += "No jobs defined`n" }
	$MenuText += "`n [0]	Return to Job Definitions menu`n"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host "Choose an item to delete or 0 to exit"
	if(Is-StringNumeric $MenuChoice)
	{ $MenuChoice = [int]$MenuChoice }
	else { return $MenuChoice }
	switch ($MenuChoice)
	{
		0 { return 0 }
		{ ($_ -ge 1) } {	# trap any number 1 or higher
			if($MenuChoice -gt $NameList.Count)
			{ return $MenuChoice }	# don't process out-of-range numbers
			$JobName = $NameList["$MenuChoice"]
			if ($JobName -ne $null)
			{
				ForEach ($JobNode in $JobNodes)
				{
					if ($JobNode.JobName -eq $JobName)
					{
						$JobNode.ParentNode.RemoveChild($JobNode)	# Yes, the object was just told to tell its parent to kill it. How cruel. A nice "Remove" option would be preferable but there isn't one.
						# break is an option; however, in the unlikely event of duplicates, keep moving
					}
				}
			}
			Save-ConfigData
			return [int]$MenuChoice
		}
		default { return $MenuChoice }
	}
}

# Menu to add new job definitions or modify existing ones
function Show-AddModifyJobMenu {
	$JobNodes = $global:ConfigData.SelectNodes("/Root/Jobs/Job")
	Show-Header "Add/Modify Export Job Definitions"
	$Base = 1		# holds the number of the job that the user will select as an option
	$NameList = @{"1"=""}		# hash table of job names
	$MenuText = " [$Base]	Add new job`n"	# Not an actual job
	ForEach ($JobNode in $JobNodes)
	{
		$Base++
		$NameList.Add("$Base",$JobNode.JobName)
		$MenuText += " [$Base]	" + $JobNode.JobName + "`n"
	}
	$MenuText += "`n [0]	Exit to job definitions menu"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host -Prompt "`nChoose an option"

	# the following checks aren't always necessary, app usually works OK without, but no one ever died from input sanitization
	if(Is-StringNumeric $MenuChoice)
	{ $MenuChoice = [int]$MenuChoice }
	else
	{ return $MenuChoice }
	switch ($MenuChoice)
	{
		0	{ return 0 }
		{($_ -ge 1) } {	# trap digits 1 or higher
			# The $_ variable is the item being checked by the switch, so the below means "find item X in the hash table and pass
			# its matching string to the function. Since it's a hash table and not an array, querying for an item that doesn't exist
			# causes nothing to happen.
			Show-AddModifyJobItem $NameList.Get_Item("$_")
			return [int]$MenuChoice
		}
		default { return -1 }	# Will cause the menu to recycle
	}
}

## The user has selected a job to modify; this menu lets it be modified
##
## $JobName = text name of the job to be modified
##		- Nothing in the code forces JobName to be unique in the XML file.
function Show-AddModifyJobItem {
	param([string]$JobName = "")
	$ConfigNode = $global:ConfigData.SelectSingleNode("Root/Configuration")
	$JobsNodes = $global:ConfigData.SelectNodes("Root/Jobs/Job")
	$ActiveJobNode = $null
	if ($JobName.Length -eq 0)
	{
		Show-Header "Create New Job"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor "Note: Entering the name of an existing job will open that job for editing.`nJob names are not case-sensitive.`n"
		$JobName = Read-Host -Prompt "Enter name of job. [ENTER] to cancel"
		if ($JobName.Length -eq 0)
		{ return }
	}
	# If the Jobs node doesn't exist, make one
	if ($global:ConfigData.SelectSingleNode("Root/Jobs") -eq $null)
	{
		$ActiveJobNode = $global:ConfigData.SelectSingleNode("/Root")
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("Jobs"))
		Save-ConfigData
		$ActiveJobNode = $null	# Used in a later test
	}
	else
	{	# placing in an else isn't strictly necessary but might save a billionth of a second in CPU cycles and that's really important (don't fall in the sarchasm)
		ForEach ($JobNode in $JobNodes)
		{
			if ($JobNode.JobName -eq $JobName)
			{
				$ActiveJobNode = $JobNode
				break	# no need to process all entries
			}
		}
	}
	if ($ActiveJobNode -eq $null)	# This would have changed if the job name had been found
	{
		# Create a node "Job" under "Jobs" and give it a "Name" attribute of whatever the user typed
		$ActiveJobNode = $global:ConfigData.SelectSingleNode("/Root/Jobs")
		$NewJobElement = $global:ConfigData.CreateElement("Job")
		$NewJobElement.SetAttribute("JobName", "$JobName")
		$ActiveJobNode = $ActiveJobNode.AppendChild($NewJobElement)
	}

	# TODO: If the script is upgraded with new options, ensure they are checked below
	# TODO: Implement better error-checking: only checks to see if the nodes are non-present or empty, not for invalid data
	# TODO: Find a way to loop these, too much almost-duplicate code

	# Check for IncludeOrExclude
	if ($ActiveJobNode.IncludeOrExclude -eq $null)
	{
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("IncludeOrExclude"))
	}
	if ($ActiveJobNode.IncludeOrExclude.Length -eq 0)
	{
		$ActiveJobNode.IncludeOrExclude = $DefaultIncludeOrExclude
	}

	# Check for ExportPath
	if ($ActiveJobNode.ExportPath -eq $null)
	{
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("ExportPath"))
	}
	if ($ActiveJobNode.ExportPath.Length -eq 0)
	{
		$ActiveJobNode.ExportPath = $DefaultExportPath
	}

	# Check for ActionBeforeExport
	if ($ActiveJobNode.ActionBeforeExport -eq $null)
	{
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("ActionBeforeExport"))
	}
	if ($ActiveJobNode.ActionBeforeExport.Length -eq 0)
	{
		$ActiveJobNode.ActionBeforeExport = $DefaultActionBeforeExport
	}

	#Check for MoveMissingVMsToThisNode
	if ($ActiveJobNode.MoveMissingVMsToThisNode -eq $null)
	{
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("MoveMissingVMsToThisNode"))
	}
	if ($ActiveJobNode.MoveMissingVMsToThisNode.Length -eq 0)
	{
		$ActiveJobNode.MoveMissingVMsToThisNode = $DefaultMoveMissingVMsToThisNode
	}

	$StayInMenu = $TRUE	# Cheap way to keep the menu displaying until it's been told to go away
	while ($StayInMenu)
	{
		if ($global:ConfigData.Root.Configuration.UseCluster -eq "No") # if cluster functions are disabled, LiveMigration isn't an option
		{	$MissingAction = "Globally Disabled" }
		else { $MissingAction = $ActiveJobNode.MoveMissingVMsToThisNode }
	 	Show-Header ("Modify Job [" + $ActiveJobNode.JobName + "]")
	 	$MenuText =
	 		" [1]	Name`n" +
	 		"	[" + $ActiveJobNode.JobName + "]`n" +
			" [2]	Export path`n" +
			"	[" + $ActiveJobNode.ExportPath + "]`n" +
			" [3]	Save or shut down VMs prior to export?`n" +
			"	[" + $ActiveJobNode.ActionBeforeExport + "]`n" +
			" [4]	Move missing VMs to this node?`n" +
			"	[$MissingAction]`n" +
			" [5]	Include or exclude selected VMs (exclude none to export all)`n" +
			"	[" + $ActiveJobNode.IncludeOrExclude + "]`n" +
			" [6]	Modify " + ($ActiveJobNode.IncludeOrExclude).ToLower() + "d VMs`n"+
			"`n [0]	Exit to job definitions menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host -Prompt "`nChoose an option"
		switch ($MenuChoice)
		{
			0	{	# Exit
				Save-ConfigData	# Commit user changes upon exit. If the user didn't change anything, commit the changes anyway.
				$StayInMenu = $FALSE
			}
			1 { # Rename the job
				Clear-Host
				$NewName = Read-Host ("Enter new job name (Enter to keep [" + $ActiveJobNode.JobName + "])")
				if ($NewName.Length)
				{
					$ActiveJobNode.JobName = "$NewName"
				}
			}
			2 { # Select the folder where this job should place its exported VMs
				if ($ActiveJobNode.ExportPath.Length -gt 0)
				{ $InitialDirectory = $ActiveJobNode.ExportPath }
				else
				{	$InitialDirectory = $ConfigNode.DefaultExportPath }
				$NewPath = Get-FolderSelection -Title ("Select Export Path for " + $ActiveJobNode.JobName) -InitialDirectory $InitialDirectory
				if ($NewPath.Length -gt 0)
				{
					if (Test-Path($NewPath))
					{
						$ActiveJobNode.ExportPath = "$NewPath"
					}
					# TODO: else { let the user know that didn't work }
				}
			}
			3 { # Toggle between "Save" and "Shutdown" for default pre-export handling of VMs
				if ($ActiveJobNode.ActionBeforeExport -eq "Save")
				{ $ActiveJobNode.ActionBeforeExport = "Shutdown" }
				else
				{ $ActiveJobNode.ActionBeforeExport = "Save" }
			}
			4 { # Should missing VMs be LiveMigrated to this node?
				if ($global:ConfigData.Root.Configuration.UseCluster -eq "Yes") # Only valid if cluster operations are enabled
				{
					if ($ActiveJobNode.MoveMissingVMsToThisNode -eq "Yes")
					{ $ActiveJobNode.MoveMissingVMsToThisNode = "No" }
					else
					{ $ActiveJobNode.MoveMissingVMsToThisNode = "Yes" }
				}
			}
			5 { # The user can select specific VMs; should those be included or excluded?
				if($ActiveJobNode.IncludeOrExclude -eq "Include")
				{	$ActiveJobNode.IncludeOrExclude = "Exclude" }
				else
				{	$ActiveJobNode.IncludeOrExclude = "Include" }
			}
			6 {	# Work with the specially defined VMs in this job
				while(Modify-SpecificVMsInJob) {}
			}
			default { }	# Will cause the menu to recycle
		}
	}
}

# Menu to work with the VMs defined within this job
function Modify-SpecificVMsInJob
{
	# TODO: Determine if it would be more efficient to use a single multi-dimensional array without overdoing complexity
	[string[]]$VMList = @()					# an array to hold all defined VMs
	[string[]]$VMListSpecial = @()	# an array to hold VMs with overrides

	# Does the VMList node exist in this job?
	if ($ActiveJobNode.SelectSingleNode("VMList") -eq $null)
	{
		# if not, make one
		$ActiveJobNode.AppendChild($global:ConfigData.CreateElement("VMList"))
	}
	$VMNodes = $ActiveJobNode.SelectNodes("VMList/VM")
	$VMCount = $VMNodes.Count
	if($VMCount)	# not much point cycling through an empty list
	{
		ForEach ($VMNode in $VMNodes)
		{
			$IsSpecial = $FALSE
			$ThisName = $VMNode.Name
			if($ActiveJobNode.IncludeOrExclude -eq "Include")	# "specials" are ignored in an Exclude job
			{
				if($VMNode.ActionBeforeExport -ne $null)
				{ $IsSpecial = $TRUE }
				if($VMNode.MoveIfMissing -ne $null)
				{	$IsSpecial = $TRUE }
			}
			if($IsSpecial)
			{ $VMListSpecial += "$ThisName" }
			else
			{ $VMList += "$ThisName" }
		}
		[string[]]$VMList = $VMList | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE}	# sort, enforce string or things are weird
		[string[]]$VMListSpecial = $VMListSpecial | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE}	# sort, enforce string or things are weird
	}
	if ($ActiveJobNode.IncludeOrExclude -eq "Include")
	{ $ShowOptionsColumns = $TRUE }
	else { $ShowOptionsColumns = $FALSE }
	Show-Header ("Work With VMs in Job [" + $ActiveJobNode.JobName + "]")
	$MenuText = ""
	$WindowWidth = $Host.UI.RawUI.WindowSize.Width - 1	# move to "constants"? or leave here on the odd chance user changes window width during operation?

	# the goal of this next section is to show all the VMs defined in this job. Since PowerShell doesn't have a built-in neat way to
	# word-wrap... well, anything... do it manually
	if($VMList.Count -ge 1)
	{
		$LinePosition = 0
		$MenuText += "Virtual Machines " + $ActiveJobNode.IncludeOrExclude + "d by this job:`n`n"
		ForEach($VM in $VMList)
		{
			if ($LinePosition -eq 0)	# this will only trigger for the very first entry
			{
				# trap the first entry so as not to generate any leading commas
				$LinePosition = 2	# the start of each line will start 2 characters from the edge...
				$MenuText += " "	# because of this
			}
			elseif (($LinePosition + $VM.Length) -gt $WindowWidth)
			{	# adding in the next name would have gone past the edge of the line....
				$MenuText += ",`n "		# so add a comma at the end of the current one, drop a line down, and...
				$LinePosition = 2			# ... move two characters over
			}
			else
			{ $MenuText += ", " }		# plenty of room on this line and it's not the first entry, so drop a comma and a space
			$MenuText += $VM
			$LinePosition += $VM.Length + 2 # +2 because: don't forget to account for the spaces and commas
		}
		$MenuText += "`n`n"
	}

	# Format the special list just like the standard list
	if($VMListSpecial.Count -ge 1)
	{
		$LinePosition = 0
		$MenuText += "Virtual Machines with Special Conditions Currently Defined in this Job:`n`n"
		ForEach($VM in $VMListSpecial)
		{
			if ($LinePosition -eq 0)	# this will only trigger for the very first entry
			{
				$LinePosition = 2
				$MenuText += " "
			}
			elseif (($LinePosition + $VM.Length) -gt $WindowWidth)
			{ $MenuText += ",`n "
				$LinePosition = 2
			}
			else
			{ $MenuText += ", " }
			$MenuText += $VM
			$LinePosition += $VM.Length + 2
		}
		$MenuText += "`n`n"
	}
	$MenuText += " [1]	Add VM`n"
	if($VMCount)
	{	# if nothing to remove, don't give the option
		$MenuText += " [2]	Remove VM`n"
		if($ActiveJobNode.IncludeOrExclude -eq "Include")
		{ $MenuText += " [3]	Set or remove special conditions for a VM`n" }
	}
	$MenuText += "`n [0]	Return to modify job menu"
	Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
	$MenuChoice = Read-Host "Choose an option"
	if (Is-StringNumeric $MenuChoice)
	{ $MenuChoice = [int]$MenuChoice }
	else
	{ return $MenuChoice }
	switch ($MenuChoice)
	{
		0 { return 0 }
		1 {
			Show-AddVM
			return 1
		}
		2 {
			if($VMCount)
			{	# those sneaky users will try to punch in hidden menu items
				Show-RemoveVMFromJob
				return 2
			}
		}
		3 {
			if($ActiveJobNode.IncludeOrExclude -eq "Include")
			{ Show-OverrideVM }	# override options are only valid in include jobs
			return 3
		}
		default { return $MenuChoice }
	}
}

# Menu to add a VM
function Show-AddVM
{
	$StayInAddMenu = 1
	$ResultBuffer = ""
	while ($StayInAddMenu -ne 0)
	{
		$MenuText = "$ResultBuffer"		# Users like feedback; not all possibilites in this menu have obvious signs
		$EligibleVMs = @("None")			# Array of discovered VMs that could be added but haven't been
		$Count = 0
		[Int]$LinePosition = 1
		# Although the variables from calling functions are visible here, it's easiest to just rebuild the list
		$VMAll = $ActiveJobNode.SelectNodes("VMList/VM")
		[String[]]$VMAllList = @()									# Array of all VMs
		[string[]]$Hosts = @()											# Container for host name(s)

		Clear-Host	# will be an interim screen prior to displaying the menu
		if ($global:ConfigData.Root.Configuration.UseCluster -eq "Yes")
		{
			try
			{
				Import-Module "FailoverClusters" -ErrorAction Stop	# Get-Cluster won't work without this
				ForEach($Server in Get-ClusterNode)
				{
					$Hosts += $Server.Name
				}
				$MenuText = " (*) Indicates a VM currently located on another host`n`n"
			}
			catch
			{
				# TODO: Let the user know that we wanted to talk to a cluster but there was an issue? Or not. Program works either way.
			}
		}
		else
		{ $Hosts += $env:computername }
		# when querying the list of VMs from a cluster, this could take a while, so let the user know
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor "Retrieving the list of available virtual machines..."

		# The next part is fairly self-explanatory -- query the host(s), look for VMs. The coding of this section turned up an oddity.
		# At a regular command-line, if you manually type in the following code and use a comma-separated list of the hosts after -ComputerName,
		# it works exactly as expected. Ex: Get-WmiObject -ComputerName vmhost1,vmhost2,vmhost3... will return all VMs on all those hosts.
		# However, if you pass in a string variable with a comma, the whole thing breaks down. Pass in an array instead.
		$VMObjects = Get-WmiObject -ComputerName $Hosts -namespace root\virtualization\v2 MSVM_ComputerSystem -filter "ElementName <> Name" | Sort-Object "ElementName"
		Show-Header ("Select VMs to Add to [" + $ActiveJobNode.JobName + "]")
		if($VMAll.Count)
		{
			ForEach ($VM in $VMAll)
			{
				$VMAllList += $VM.Name
			}
			$VMAllList = $VMAllList | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE}
		}

		# the following section formats the list of available VMs for the user to choose from
		ForEach ($VMObject in $VMObjects)
		{	# VMObjects is an array of WMI objects with predefined element names
			if ($VMAllList -contains $VMObject.ElementName)
			{ continue }	# If the list of defined VMs contains this VM, skip it
			$Count++
			$EligibleVMs += $VMObject.ElementName
			$DisplayName = $VMObject.ElementName
			if($VMObject.__SERVER -ne $env:computername)
			{ $DisplayName = "* " + $DisplayName }
			$MenuText += ((" [" + $Count + "]").PadRight(7) + $DisplayName).PadRight(25)
			if($LinePosition -eq $ObjectsPerLine)
			{
				$LinePosition = 1
				$MenuText += "`n"
			}
			else
			{
				$LinePosition++
				$MenuText += " "
			}
		}
		if($Count)
		{ $MenuText += "`n`n [-1] To add all" }
		else
		{ $MenuText += "No available VMs" }
		$MenuText += "`n [0] Exit to VM menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host "Choose an option OR type a name (accepts unlisted VMs)"
		if (Is-StringNumeric $MenuChoice)
		{
			$MenuChoice = [int]$MenuChoice	# explicit cast necessary for following comparators to work as expected
			switch ($MenuChoice)
			{
				-1 {
					for($i = 1; $i -lt $EligibleVMs.Count; $i++)
					{
						Add-VMToJob $EligibleVMs[$i]
					}
				}
				{ $_ -ge 1 } {
					if($MenuChoice -gt ($EligibleVMs.Count - 1))
					{ $ResultBuffer = "Invalid numeric entry [$MenuChoice]`n" }
					else
					{
						$ResultBuffer = Add-VMToJob $EligibleVMs[[int]$MenuChoice]
					}
				}
			}
		}
		else
		{	# If the user keys in a non-numeric, then take that input and put it in the list of VMs
			$MenuChoice = [String]$MenuChoice
			if (($VMList -notcontains $MenuChoice) -and ($VMListSpecial -notcontains $MenuChoice))
			{
				$ResultBuffer = Add-VMToJob $MenuChoice
			}
			else
			{ $ResultBuffer = "VM [$MenuChoice] is already listed`n" }
		}
		$StayInAddMenu = $MenuChoice
	}
}

## Actual behind the scenes function that will add the VM to the job
##
## $VMToAdd = the text name of the job to be added
function Add-VMToJob
{
	param([String]$VMToAdd)
	$ListNode = $ActiveJobNode.SelectSingleNode("VMList")
	$NewVMElement = $global:ConfigData.CreateElement("VM")
	$NewVMElement.SetAttribute("Name", "$VMToAdd")
	$ListNode.AppendChild($NewVMElement) | Out-Null
	Save-ConfigData
	return "Added VM [$VMToAdd]`n`n"
}

## Present a menu of defined VMs that can be removed from a job
function Show-RemoveVMFromJob
{
	do
	{
		Show-Header ("Remove VMs from [" + $ActiveJobNode.JobName + "]")
		$Count = 0
		$MenuText = ""
		# Easiest to just rebuild the list
		# $ActiveJobNode variable is visible here from calling functions. The variable itself is read-only at this level,
		# but the underlying XML that it contains can be manipulated.
		$VMAll = $ActiveJobNode.SelectNodes("VMList/VM")
		[string[]]$VMAllList = @()
		if($VMAll.Count)
		{
			ForEach ($VM in $VMAll)
			{
				$VMAllList += $VM.Name
			}
			$VMAllList = $VMAllList | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE}
			$LinePosition = 1
			ForEach($VM in $VMAllList)
			{
				$Count++
				$MenuText += ((" [" + $Count + "]").PadRight(7) + $VM).PadRight(24)
				if($LinePosition -eq $ObjectsPerLine)
				{
					$LinePosition = 1
					$MenuText += "`n"
				}
				else
				{
					$LinePosition++
					$MenuText += " "
				}
			}
			$MenuText += "`n`n [-1] to remove all"
		}
		else
		{ $MenuText += "No VMs defined in this job`n" }
		$MenuText += "`n [0] to return to VM menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host "Choose an option"
		if (Is-StringNumeric $MenuChoice)
		{
			$MenuChoice = [int]$MenuChoice
			if ($MenuChoice -eq -1)
			{
				ForEach ($VM in $VMAllList)
				{
					Remove-VMFromJob $VM
				}
			}
			if(($MenuChoice -gt 0) -and ($MenuChoice -le $VMAllList.Count))
			{
				Remove-VMFromJob $VMAllList[($MenuChoice - 1)]
			}
		}
	} until ($MenuChoice -eq 0)
}

## This is the function that actually removes VMs from the job definition
##
## $VMToRemove = text name of the job to be removed
function Remove-VMFromJob
{
	param([String]$VMToRemove)
	# $VMNodes variable is visible here from calling functions. The variable itself is read-only at this level,
	# but the underlying XML that it contains can be manipulated.
	ForEach ($VMNode in $VMNodes)
	{
		if($VMNode.Name -eq $VMToRemove)
		{
			$VMNode.ParentNode.RemoveChild($VMNode)	#Yes, again, a "Kill me, Mom" function
			Save-ConfigData
			# could break out and return, but this will mop up any duplicates that somehow slipped in
		}
	}
}

## Menu that allows user to override default settings for VMs in an include job
function Show-OverrideVM
{
	$StayInOverrideMenu = 1
	while($StayInOverrideMenu -ne 0)
	{
		if ($VMNodes.Count)
		{
			[String[]]$VMListExportOverride = @()		# array for VMs that override pre-export behavior
			[String[]]$VMListMoveOverride = @()			# array for VMs that override LiveMigration behavior
			ForEach($VMNode in $VMNodes)		# VMNodes visible from calling functions
			{
				if($VMNode.ActionBeforeExport -ne $null)
				{ $VMListExportOverride += $VMNode.Name }
				if($VMNode.MoveIfMissing -ne $null)
				{	$VMListMoveOverride += $VMNode.Name }
			}

			# For the first part of the menu, let the user see what VMs have overrides. The overrides themselves are not visible
			# simply because of the logistical challenge of showing a potentially large list with multiple options in readable formats
			# on a 80x25 character screen
			Show-Header ("Override Defaults for VMs in [" + $ActiveJobNode.JobName + "]")
			$MenuText = ""
			$WindowWidth = $Host.UI.RawUI.WindowSize.Width - 1
			if($VMListExportOverride.Count -ge 1)
			{
				$MenuText += "Virtual Machines that override save/shutdown settings`n (Overrides do NOT change if job default is changed):`n`n"
				$LinePosition = 0
				ForEach($VM in $VMListExportOverride)
				{
					if ($LinePosition -eq 0)	# this will only trigger for the very first entry
					{
						$LinePosition = 2
						$MenuText += " "
					}
					elseif (($LinePosition + $VM.Length) -gt $WindowWidth)
					{ $MenuText += ",`n "
						$LinePosition = 2
					}
					else
					{ $MenuText += ", " }
					$MenuText += $VM
					$LinePosition += $VM.Length + 2
				}
				$MenuText += "`n`n"
			}
			else { $MenuText += "All VMs in this job use default setting of " + $ActiveJobNode.ActionBeforeExport.ToLower() + " before export`n`n" }

			if($VMListMoveOverride.Count -ge 1)
			{
				$LinePosition = 0
				$MenuText += "Virtual Machines that override default move setting`n (Overrides do not change if job default is changed):`n`n"
				ForEach($VM in $VMListMoveOverride)
				{
					if ($LinePosition -eq 0)	# this will only trigger for the very first entry
					{
						$LinePosition = 2
						$MenuText += " "
					}
					elseif (($LinePosition + $VM.Length) -gt $WindowWidth)
					{ $MenuText += ",`n "
						$LinePosition = 2
					}
					else
					{ $MenuText += ", " }
					$MenuText += $VM
					$LinePosition += $VM.Length + 2
				}
				$MenuText += "`n`n"
			}
			else
			{
				if($ActiveJobNode.MoveMissingVMsToThisNode -eq "No")
				{ $nottext = "not " }
				else
				{ $nottext = "" }
				$MenuText += "All VMs in this job use default setting to " + $nottext + "move to this node before export`n`n"
			}
			$MenuText += " [1]	Override save or shutdown behavior on individual VMs`n" +
				" [2]	Override migration behavior on individual VMs`n"
		}
		else { $MenuText = "There are no VMs defined in this job.`n" } #this should not be a reachable message, but just in case
		$MenuText += "`n [0]	Return to job menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host "`nChoose an option"
		if(Is-StringNumeric $MenuChoice)
		{
			$MenuChoice = [int]$MenuChoice
			switch ($MenuChoice)
			{
				1 { Show-OverrideSaveShutdown	}
				2 { Show-OverrideMoveMissingVM }
			}
		}
		$StayInOverrideMenu = $MenuChoice
	}
}

## This menu allows the user to select VMs that will override default save/shutdown behavior
function Show-OverrideSaveShutdown
{
	do
	{
		Show-Header ("Override Save/Shutdown Behavior in [" + $ActiveJobNode.JobName + "]")
		$MenuText = ""
		$VMAll = $ActiveJobNode.SelectNodes("VMList/VM")
		$VMUnsortedList = @()
		$VMAllList = @()
		$VMAllList += ,@("?FakeSeedVM?", "Default") # this gets around oddness with automatic multi-dimensional arrays that only have one sub-array
		if($VMAll.Count)
		{
			ForEach ($VM in $VMAll)
			{
				if($VM.ActionBeforeExport -ne $null)
				{ $option = $VM.ActionBeforeExport }
				else
				{ $option = "Default" }
				$VMUnsortedList += ,@($VM.Name, $Option)
			}
			if($VMUnsortedList.Count -gt 1)	# Again, mostly for oddness with multi-dimensional arrays
			{ $VMAllList += $VMUnsortedList | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE} }
			else
			{ $VMAllList += $VMUnsortedList }
			$LinePosition = 1
			$Count = 0
			ForEach($VM in $VMAllList)
			{
				if ($VM[0] -eq "?FakeSeedVM?") # this record only exists to make the array work; skip it
				{ continue }
				$Count++
				$MenuText += ((" [$Count] ").PadRight(6) + ("(" + $VM[1] + ")").PadRight(11) + $VM[0]).PadRight(38)
				if($LinePosition -eq $LargeObjectsPerLine)
				{
					$LinePosition = 1
					$MenuText += "`n"
				}
				else
				{
					$LinePosition++
					$MenuText += " "
				}
			}
			$MenuText += "`n`n Enter the item number to toggle it between Save, Shutdown, and Default`n" +
				"	-`"Save`" will place the VM in a saved state prior to export`n" +
				"	-`"Shutdown`" shuts the VM down (use for domain controllers)`n" +
				"	-`"Default`" follows the default for this job (currently [" + $ActiveJobNode.ActionBeforeExport + "])`n" +
				"	Note: ALL VMs are returned to their prior power state after export`n"
		}
		else
		{ $MenuText += "No VMs defined in this job`n" }	# should never be reachable
		$MenuText += "`n [0] to return to VM menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host "Choose an option"
		if (Is-StringNumeric $MenuChoice)
		{
			$MenuChoice = [int]$MenuChoice
			if(($MenuChoice -gt 0) -and ($MenuChoice -le $VMAllList.Count))
			{
				# First, figure out what the next step in the cycle is
				switch ([String]($VMAllList[($MenuChoice)][1]))
				{
					"Default" { $NewOption = "Save"	}
					"Save" { $NewOption = "Shutdown" }
					"Shutdown" { $NewOption = "Default" }
					default { $NewOption = "Default" }
				}
				ForEach ($VM in $VMAll)
				{
					if ($VM.Name -eq $VMAllList[($MenuChoice)][0])
					{
						$NodeToManipulate = $VM.SelectSingleNode("ActionBeforeExport")
						if ($NewOption -eq "Default")
						{
							$NodeToManipulate.ParentNode.RemoveChild($NodeToManipulate)
						}
						else
						{
							if($NodeToManipulate -eq $null)
							{
								$NewElement = $global:ConfigData.CreateElement("ActionBeforeExport")
								$NewElement.InnerText = $NewOption
								$VM.AppendChild($NewElement)
							}
							else { ($NodeToManipulate.InnerText = $NewOption) }
						}
						break
					}
				}
			}
		}
	} until ($MenuChoice -eq 0)
	Save-ConfigData
}

## This menu allows the user to select VMs that will override default LiveMigration behavior
## TODO: This is almost an exact duplicate of the code for the pre-export action. Find a way to collapse them.
function Show-OverrideMoveMissingVM
{
	do
	{
		Show-Header ("Override Missing VM Behavior in [" + $ActiveJobNode.JobName + "]")
		if ($global:ConfigData.Root.Configuration.UseCluster -eq "No")
		{ $MenuText = " -- Cluster operations are globally disabled; these settings will be IGNORED --`n`n" }
		else { $MenuText = "" }
		$VMAll = $ActiveJobNode.SelectNodes("VMList/VM")
		$VMUnsortedList = @()
		$VMAllList = @()
		$VMAllList += ,@("?FakeSeedVM?", "Default") # this gets around oddness with automatic multi-dimensional arrays that only have one sub-array
		if($VMAll.Count)
		{
			ForEach ($VM in $VMAll)
			{
				if($VM.MoveIfMissing -ne $null)
				{ $option = $VM.MoveIfMissing }
				else
				{ $option = "Default" }
				$VMUnsortedList += ,@($VM.Name, $Option)
			}
			if($VMUnsortedList.Count -gt 1)	# Again, mostly for oddness with multi-dimensional arrays
			{ $VMAllList += $VMUnsortedList | Sort-Object @{Expression={$_[0]}; Ascending=$TRUE} }
			else
			{ $VMAllList += $VMUnsortedList }
			$LinePosition = 1
			$Count = 0
			ForEach($VM in $VMAllList)
			{
				if ($VM[0] -eq "?FakeSeedVM?") # this record only exists to make the array work; skip it
				{ continue }
				$Count++
				$MenuText += ((" [$Count] ").PadRight(6) + ("(" + $VM[1] + ")").PadRight(11) + $VM[0]).PadRight(38)
				if($LinePosition -eq $LargeObjectsPerLine)
				{
					$LinePosition = 1
					$MenuText += "`n"
				}
				else
				{
					$LinePosition++
					$MenuText += " "
				}
			}
			$MenuText += "`n`n Enter the item number to toggle it between Yes, No, and Default`n" +
				"	-`"Yes`" attempts to migrate the VM to this node prior to export`n" +
				"	-`"No`" skips the VM if it is not on this node`n" +
				"	-`"Default`" follows the job setting (currently [" + $ActiveJobNode.MoveMissingVMsToThisNode + "])`n"
		}
		else
		{ $MenuText += "No VMs defined in this job`n" }	# should never be reachable
		$MenuText += "`n [0] to return to VM menu"
		Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor $MenuText
		$MenuChoice = Read-Host "Choose an option"
		if (Is-StringNumeric $MenuChoice)
		{
			$MenuChoice = [int]$MenuChoice
			if(($MenuChoice -gt 0) -and ($MenuChoice -le $VMAllList.Count))
			{
				# First, figure out what the next step in the cycle is
				switch ([String]($VMAllList[($MenuChoice)][1]))
				{
					"Default" { $NewOption = "Yes"	}
					"Yes" { $NewOption = "No" }
					"No" { $NewOption = "Default" }
					default { $NewOption = "Default" }
				}
				ForEach ($VM in $VMAll)
				{
					if ($VM.Name -eq $VMAllList[($MenuChoice)][0])
					{
						$NodeToManipulate = $VM.SelectSingleNode("MoveIfMissing")
						if ($NewOption -eq "Default")
						{
							$NodeToManipulate.ParentNode.RemoveChild($NodeToManipulate)
						}
						else
						{
							if($NodeToManipulate -eq $null)
							{
								$NewElement = $global:ConfigData.CreateElement("MoveIfMissing")
								$NewElement.InnerText = $NewOption
								$VM.AppendChild($NewElement)
							}
							else { ($NodeToManipulate.InnerText = $NewOption) }
						}
						break
					}
				}
			}
		}
	} until ($MenuChoice -eq 0)
	Save-ConfigData
}

############ End Functions ############

############# Main Routine ############

do {} until (Read-ConfigData)	# Read configuration -- problems in that routine will force an Exit
while (Show-MainMenu) {}			# Loop the main menu until user exits
Clear-Host										# Wipe the screen before exit

########## End Main Routine ###########