<#
 Export-VMs.ps1 Version 1.4
 Export VMs on a Hyper-V Server
 
 Corrected MSVM Classes Namespace path to root\virtualization\v2
 Removed ExportVirtualSystemEx Method since it is obsolete
 Added ExportSystemDefinition Method to Export VMs
 Removed Cluster Checks and Migration Options
 Updated Hosts and VM Node Checks within Array
#>

## command-line parameters must be trapped immediately or they'll throw an error
## $JobName = the name of the job (defined in VMExportConfig.xml) to be run; if none, an "export-all with defaults" job is executed
## $Subfolder = the name of a subfolder to be used/created under the export path; if empty, use the folder as defined
param ([String]$JobName="",[String]$Subfolder="")			# this must be trapped right at the top; if no job is specified, just export everything

## Global "Constants" (actually variables, but are unchangeable by functions)
$ScriptDir = Split-Path (Get-Variable MyInvocation).Value.MyCommand.Path
$ConfigFile = "$ScriptDir\VMExportConfig.xml"				# Modify this to change where configuration data is saved
$DefaultLogFile = "$ScriptDir\VMExport.log"					# Modify this to change default log location (overriden when config file is read)
$OpsTimeOut = 360											# Seconds to wait for a process before continuing without it (mostly, how long to wait for a VM to finish changing state before skipping it)

## Global variables
$global:ConfigData = ""										# Holds configuration data as read from the config XML file
[String]$global:LogFile	= $DefaultLogFile					# Where messages will be stored
$ThisHost = Get-Content env:computername					# Get the name of this computer
$global:ActionBeforeExport = "Save"							# Will be reset by config load
$global:IncludeOrExclude	= "Exclude"						# Will be reset by config load
[String[]]$global:Hosts = @()								# Array of host names

############## Functions ##############

## Log errors with timestamp
function LogMessage
{
	param([String]$Condition,[String]$ErrorMessage)
	$Stamp = Get-Date
	Add-Content $global:LogFile "$Stamp $Condition`: $ErrorMessage"
}

## Read the configuration file and ensure it is more or less valid format
function Read-ConfigData
{
	$global:ConfigData = New-Object XML
	if (Test-Path $ConfigFile)
	{
		try
		{
			$Error.Clear()
			$global:ConfigData.Load($ConfigFile)
			return $TRUE
		}
		catch
		{
			Clear-Host
			LogMessage "Error" ("Error reading configuration file: $Error.")
			return $FALSE
		}
	}
	else
	{
		LogMessage "Error" ("Configuration file not found: " + $ConfigFile)
		return $FALSE
	}
	$ConfigurationRoot = $global:ConfigData.Root.Configuration

	if ($ConfigurationRoot -eq $null)
	{
		LogMessage "Error" "Configuration node missing in " + $ConfigFile
		return $FALSE
	}

	# Check to be sure the DefaultExportPath key exists and is set
	if ($ConfigurationRoot.DefaultExportPath -eq $null)
	{
		LogMessage "Error" "`"DefaultExportPath`" not found in configuration file."
		return $FALSE
	}
	if ($ConfigurationRoot.DefaultExportPath.Length -eq 0)
	{
		LogMessage "Error" "`"DefaultExportPath`" not set in configuration file."
		return $FALSE
	}
	# Check to be sure the LogFile key exists and is set
	if ($ConfigurationRoot.LogFile -eq $null)
	{
		LogMessage "Error" "`"LogFile`" not found in configuration file."
		return $FALSE
	}
	if ($ConfigurationRoot.LogFile.Length -eq 0)
	{
		LogMessage "Error" "`"LogFile`" not set in configuration file."
		return $FALSE
	}
	# Check to be sure the DefaultActionBeforeExport key exists and is set
	if ($ConfigurationRoot.DefaultActionBeforeExport -eq $null)
	{
		LogMessage "Error" "`"DefaultActionBeforeExport`" not found in configuration file."
		return $FALSE
	}
	if ($ConfigurationRoot.DefaultActionBeforeExport.Length -eq 0)
	{
		LogMessage "Error" "`"DefaultActionBeforeExport`" not set in configuration file."
		return $FALSE
	}
	# Check to be sure the DefaultIncludeOrExclude key exists and is set
	if ($ConfigurationRoot.DefaultIncludeOrExclude -eq $null)
	{
		LogMessage "Error" "`"DefaultIncludeOrExclude`" not found in configuration file."
		return $FALSE
	}
	if ($ConfigurationRoot.DefaultIncludeOrExclude.Length -eq 0)
	{
		LogMessage "Error" "`"DefaultIncludeOrExclude`" not set in configuration file."
		return $FALSE
	}
}

## Import dependencies and populate globals
function SetUp-Environment
{
	$global:LogFile = $global:ConfigData.Root.Configuration.LogFile
	$global:ActionBeforeExport = $global:ConfigData.Root.Configuration.DefaultActionBeforeExport
}

## $VMName = string name of the VM to be exported
## $Destination = string indicating fully qualified destination path including any custom Subfolder
## $SaveOrShutDown = if powered on, should this VM be saved or shut down? [String] "Save" or "Shutdown"
function Export-SingleVM
{
	param([String]$VMName,[String]$Destination,[String]$SaveOrShutDown)

	$StartingState = 0	# Status of the VM before export begins
	$StartingHost = $env:computername
	LogMessage "Information" "Beginning export of $VMName"
	$HVMgmtService = Get-WMIObject -namespace root\virtualization\v2 MSVM_VirtualSystemManagementService

	# Locate the requested VM
	$VMObject = Get-WmiObject -ComputerName $env:computername -namespace root\virtualization\v2 -Query "SELECT * FROM MSVM_ComputerSystem WHERE ElementName = '$VMName'"
	if($VMObject -eq $null)
	{	# this means the named VM wasn't found on the local system
		LogMessage "Skipped" "VM $VMName selected for export but was not found."
		return
	}
	# The VM is selected, options are set, destination is known. Proceed with the export.

	try
	{
		$StayInLoop = $TRUE
		$FirstEntry = $TRUE
		while ($StayInLoop)
		{
			# VMObject.EnabledState doesn't refresh; retrieve the object again
			$VMObject = Get-WmiObject -ComputerName $env:computername -namespace root\virtualization\v2 -Query "SELECT * FROM MSVM_ComputerSystem WHERE ElementName = '$VMName'"
			$StartingState = $VMObject.EnabledState
			switch ($StartingState)
			{
				32768 # Paused: Do not interfere. Report and return.
				{
					LogMessage "Skipped" "VM $VMName is paused and cannot be exported."
					return
				}
				{ $_ -ge 32770 } # Wait for Timeout and Check again
				{
					if ($FirstEntry)
					{
						Start-Sleep $OpsTimeOut
					}
					else
					{
						if($StartingState -ge 32770)
						{
							LogMessage "Skipped" "VM $VMName was in a transitional state ($StartingState) when the export started and exceeded the timeout of $OpsTimeOut seconds."
							return
						}
					}
				}
				Default
				{
					$StayInLoop = $FALSE # Store State and Continue
				}
			}
		}

		if ($StartingState -eq 2)	# The VM is running, so perform the user-selected action
		{
			if($SaveOrShutDown -eq "Save")
			{
				$SaveStatus = $VMObject.RequestStateChange(32769)
				if($SaveStatus.ReturnValue -eq 4096)
				{
					$Job = [WMI]$SaveStatus.Job
					while($Job.JobState -eq 4)	# 4 is "running"
					{
						Start-Sleep 5		# wait about 5 seconds between checks
						$Job.PSBase.Get()
					}
					if($Job.JobState -ne 7)	# 7 is "completed"
					{
						LogMessage "Error" ("Attempted to save $VMName for export but encountered error " + $Job.ErrorDescription)
						return
					}
				}
			}
			else
			{
				$ShutdownObject = Get-WmiObject -namespace root\virtualization\v2 -Query "Associators of {$VMObject} WHERE AssocClass=MSVM_SystemDevice ResultClass=MSVM_ShutdownComponent"
				$ShutdownStatus = $ShutdownObject.InitiateShutdown($True, "Scheduled Export")
				if($ShutdownStatus.ReturnValue -eq 4096)
				{
					$Job = [WMI]$ShutdownStatus.Job
					while($Job.JobState -eq 4)	# 4 is "running"
					{
						Start-Sleep 
						$Job.PSBase.Get()
					}
					if($Job.JobState -ne 7)	# 7 is "completed"
					{
						LogMessage "Error" ("Requested VM $VMName to shut down, error received: " + $Job.ErrorDescription)
						return
					}
				}
				elseif ($ShutdownStatus.ReturnValue -eq 0)
				{	# Check for Shutdown
					$CheckInterval = 5	# Sleep Check Intervals in Seconds
					for($i = 0;$i -lt $OpsTimeOut;$i += $CheckInterval)
					{
						Start-Sleep $CheckInterval
						# VMObject.EnabledState doesn't refresh; call again
						$VMObject = Get-WmiObject -ComputerName $env:computername -namespace root\virtualization\v2 -Query "SELECT * FROM MSVM_ComputerSystem WHERE ElementName = '$VMName'"
						if($VMObject.EnabledState -eq 3)
						{ break }
					}
					if($VMObject.EnabledState -ne 3)
					{
						LogMessage "Error" "Attempted to shutdown $VMName for export, but timeout exceeded without a successful shutdown. Please check the status of this VM immediately."
						return
					}
				}
				else
				{
					LogMessage "Error" ("Attempted to shutdown $VMName for export, but it failed with an unexpected status: " + $ShutdownStatus.ReturnValue)
					return
				}
			}
		}

		# Check for earlier exports of this VM.
		if(Test-Path($Destination + "\" + $VMName))
		{
			# Export will attempt to create a subfolder with the name of the VM in the destination location.
			# If it already exists, wipe the destination
			Remove-Item ($Destination + "\" + $VMName) -Force -Recurse
		}
		
		# Obsolete
		# $ExportStatus = $HVMgmtService.ExportVirtualSystemEx($VMObject.__PATH, $True, $Destination) # which VM to export, copy the entire state (VHDs, AVHDs, saved state), and where to put it all
		
		# New Export Configuration
		$es = @($VMObject.GetRelated('Msvm_VirtualSystemExportSettingData'))[0] # must retype ManagementObjectCollection to array as it does not define enumerator to be indexed directly
		$es.CopySnapshotConfiguration = 0  # No Snapshots
		$es.CopyVmStorage = $true  # Include VHDs
		$es.CopyVmRuntimeInformation = $false  # No, Save state files
		$es.CreateVmExportSubdirectory = $true  # yes, create the $VMName folder under $Destination
		
		$ExportStatus = $HVMgmtService.ExportSystemDefinition($VMObject.Path.Path, $Destination, $es.GetText(1))
		if ($ExportStatus.ReturnValue -eq 4096)	# 4096 is "asynchronous job started"
		{
			$Job = [WMI]$ExportStatus.Job
			while($Job.JobState -eq 4)	# 4 is "running"
			{
				Start-Sleep 5
				$Job.PSBase.Get()
			}
			if($Job.JobState -eq 7)	# 7 is "completed"
			{ LogMessage "Information" "VM $VMName exported successfully." }
			else
			{ LogMessage "Error" ("Export of $VMName failed: " + $Job.ErrorDescription) }
		}
		else
		{
			switch ($ExportStatus.ReturnValue)
			{
				32775
				{
					LogMessage "Error" "The VM $VMName could not be exported because it is turned on and attempts to turn it off did not succeed."
					return
				}
				Default
				{ LogMessage "Error" ("Unable to start export for $VMName for an undetermined reason. Exit code " + $ExportStatus.ReturnValue) }
			}
		}
	}
	catch
	{
		return
	}
	finally
	{
		# Get the current state of the object
		$VMObject = Get-WmiObject -ComputerName $env:computername -namespace root\virtualization\v2 -Query "SELECT * FROM MSVM_ComputerSystem WHERE ElementName = '$VMName'"
		if($VMObject.EnabledState -ne $StartingState)
		{
			$ReturnStatus = $VMObject.RequestStateChange($StartingState)
			if($ReturnStatus.ReturnValue -eq 4096)
			{
				$Job = [WMI]$ReturnStatus.Job
				while($Job.JobState -eq 4)
				{
					Start-Sleep 5
					$Job.PSBase.Get()
				}
				if($Job.JobState -ne 7)
				{
					LogMessage "Error" ("Attempted to restore $VMName to its original state but encountered error " + $Job.ErrorDescription)
				}
			}
		}
	}
}

function Execute-Export
{
	$VMList = @()
	$JobExportPath = $global:DefaultExportPath
	$JobActionBeforeExport = $global:ActionBeforeExport
	$JobIncludeOrExclude = $global:DefaultIncludeOrExclude

	if($JobName.Length)
	{ LogMessage "Information" "$JobName started" }
	else
	{ LogMessage "Information" "Exporting all VMs" }

	# Build the list of VMs to be exported and the options to be used on them
	if($JobName.Length)
	{
		$JobsNode = $global:ConfigData.SelectNodes("Root/Jobs/Job")
		ForEach ($JobNode in $JobsNode)
		{
			if($JobNode.JobName -eq $JobName)
			{
				$ActiveJobNode = $JobNode
				break
			}
		}
		if($ActiveJobNode -eq $null)
		{
			# If Job is non-existent, log it as an error and exit.
			LogMessage "Error" "Unable to find specified job `"$JobName`""
			Exit
		}
		$JobExportPath = $ActiveJobNode.ExportPath
		$JobActionBeforeExport = $ActiveJobNode.ActionBeforeExport
		$JobIncludeOrExclude = $ActiveJobNode.IncludeOrExclude
		$VMNodes = $ActiveJobNode.SelectNodes("VMList/VM")
	}
	else
	{
		$JobExportPath = $global:DefaultExportPath
		$JobActionBeforeExport = $global:DefaultActionBeforeExport
	}
	if($JobName.Length -and ($JobIncludeOrExclude -eq "Include"))
	{
		# Only export from the user-defined list
		ForEach ($VMObject in $VMNodes)
		{
			if($VMObject.ActionBeforeExport -ne $null)
			{ $VMActionBeforeExport = $VMObject.ActionBeforeExport }
			else
			{ $VMActionBeforeExport = $JobActionBeforeExport }
			$VMList += ,@($VMObject.Name + ",$VMActionBeforeExport")
		}
	}
	elseif (($JobName.Length -and ($JobIncludeOrExclude -eq "Exclude")) -xor ($JobName.Length -eq 0))
	{
		[String[]]$VMExcludeList = @()				# Will hold the excluded items
		[String[]]$Hosts	= @()					# Will hold the host name(s)
		$VMActionBeforeExport = ""

		# Get a list of items to Exclude
		if($JobName.Length)
		{
			ForEach($VMObject in $VMNodes)
			{	$VMExcludeList += $VMObject.Name }
			$Hosts += $env:computername
		}
		else { $Hosts += $env:computername }

		# Retrieve the list of VMs on this server or cluster
		$VMObjects = Get-WmiObject -ComputerName $Hosts -namespace root\virtualization\v2 MSVM_ComputerSystem -filter "ElementName <> Name"

		# Step through the retrieved VMs; skip if excluded, add to export list with job settings
		ForEach ($VMObject in $VMObjects)
		{
			if ($VMExcludeList -notcontains $VMObject.ElementName)
			{
				$Count++
				$VMList += ,@($VMObject.ElementName + ",$JobActionBeforeExport")
			}
		}
	}
	else
	{
		LogMessage "Error" "Invalid value $JobIncludeOrExclude specified for job " + $JobName + "; expected `"Include`" or `"Exclude`" for IncludeOrExclude element."
		Exit
	}

	# The list of VMs to export is now prepared. The next thing to look at is the destination.

	# If data was submitted for parameter $Subfolder, modify the export path to include it
	if($Subfolder.Length)
	{
		if(($JobExportPath[($ExportPath.Length - 1)] -ne "\") -and ($Subfolder[0] -ne "\"))
		{	$JobExportPath += "\" }	# Make sure the original folder and specified subfolder are separated by a \
		$JobExportPath += $Subfolder
	}

	if(!(Test-Path($JobExportPath)))
	{
		try
		{	# If the folder doesn't exist, make it
			New-Item $JobExportPath -Type directory -Force -ErrorAction Stop | Out-Null
		}
		catch
		{
			LogMessage "Error" "Path specified for job $JobName ($JobExportPath) is not valid and could not be created. Message from system: $error"
			Exit
		}
	}

	# Step through the VMs and export them.
	ForEach ($VM in $VMList)
	{
		$VMArray = [RegEx]::Split($VM, ",")
		Export-SingleVM $VMArray[0] $JobExportPath $VMArray[1] $VMArray[2] | Out-Null
	}
}

############ End Functions ############

############# Main Routine ############

$Error.Clear()
if(!(Read-ConfigData)) { Exit }	# Retrieve the configuration or Exit
Write-Host ("Export-VMs.ps1 Version 1.1`n`n")
Write-Host ("View the log file for details: " + $global:LogFile)
SetUp-Environment
Execute-Export

########## End Main Routine ###########