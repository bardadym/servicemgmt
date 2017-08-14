# Make a list of Windows services in file Services.csv on servers specified in file Servers.csv
# 
# 2017 Viktor Bardadym

#**************************************************************************************
# Constants
#**************************************************************************************

#**************************************************************************************
# Global variables
#**************************************************************************************

#**************************************************************************************
# Functions
#**************************************************************************************

#<summary>
# Returns folder where the running script is located
#</summary>
function Get-ScriptFolder 
{ 
	Split-Path -parent $PSCommandPath 
}

#<summary>
# Adds one row per object / user to the CSV report file
#</summary>
#<param name="$server">Server name.</param>
#<param name="$serviceName">Service name.</param>
#<param name="$serviceDisplayName">Service display name.</param>
#<param name="$status">Status of the service.</param>
#<param name="$startupMode">Startup mode of the service.</param>
#<param name="$dependentServices">Space-delimited list of services which require the current service to run.</param>
#<param name="$requiredServices">Space-delimited list of services which are required by the current service to run.</param>
#<param name="$timeout">Timeout to check whether the service has been started.</param>
function Add-Row([string]$server, [string]$serviceName, [string]$serviceDisplayName, [string]$status, [string]$startupMode, [string]$dependentServices, [string]$requiredServices, [string]$timeout) 
{
	$fileRow = $server + $columnDelimitor + $serviceName + $columnDelimitor + $serviceDisplayName + $columnDelimitor + $status + $columnDelimitor + $startupMode + $columnDelimitor + $dependentServices + $columnDelimitor + $requiredServices + $columnDelimitor + $timeout
	Add-Content -Path $serviceStatusFile -Value $fileRow -Encoding UTF8
}

#<summary>
# Get service startup mode
#</summary>
#<param name="$server">Server name.</param>
#<param name="$serviceName">Service name.</param>
function Get-StartupMode([string]$server, [string]$serviceName) 
{
	$serviceNameFilter = "Name='" + $serviceName + "'"
	$serviceWmi = Get-WmiObject -Class Win32_Service -Property StartMode -Filter $serviceNameFilter -ComputerName $server
	$serviceWmi.StartMode
}

#**************************************************************************************
# Main script
#**************************************************************************************

# Assignment of constant values
$columnDelimitor = ","
$scriptFolder = Get-ScriptFolder
$serverFile = $scriptFolder + "\Servers.csv"
$serviceStatusFile = $scriptFolder + "\ServicesFullList.csv"

if (Test-Path $serviceStatusFile)
{
	#File exists; stop the script
	Write-Host "Status file already exists at" $serviceStatusFile ". Script is stopped to avoid overwriting." -ForegroundColor Red
}
else
{
	Add-Row "Server" "Service" "Description" "Status" "StartupType" "DependentServices" "RequiredServices" "Timeout"   
	$input = Import-Csv -path $serverFile
	# Iterate through service list file
	ForEach ($row in $input)
	{
		[string]$serverName = $row.Server
		$services = Get-Service "*" -ComputerName $serverName
		foreach ($service in $services)
		{
			$startupMode = Get-StartupMode $serverName $service.Name
			Write-Host $serverName ":" $service.DisplayName
			Add-Row $serverName $service.Name $service.DisplayName $service.Status $startupMode $service.DependentServices $service.RequiredServices 0
		}
	}
	Write-Host "The list of services is built in" $serviceStatusFile -ForegroundColor Green
}
