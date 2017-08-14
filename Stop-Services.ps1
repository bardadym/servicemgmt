# Stop Windows services listed in the input file Services.csv
# 
# 2017 Viktor Bardadym

#**************************************************************************************
# Constants
#**************************************************************************************

#**************************************************************************************
# Global variables
#**************************************************************************************

$Global:lineCount = 0
$Global:errorCount = 0
$Global:defaultTimeout = 10

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
function Add-Row([string]$lineCount, [string]$server, [string]$serviceName, [string]$serviceDisplayName, [string]$status, [string]$startupMode, [string]$timeout) 
{
	$fileRow = $lineCount + $columnDelimitor + $server + $columnDelimitor + $serviceName + $columnDelimitor + $serviceDisplayName + $columnDelimitor + $status + $columnDelimitor + $startupMode + $columnDelimitor + $timeout
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

#<summary>
# Recursively stops dependencies of the service, then makes the record on status and stops the service
#</summary>
#<param name="$server">Server name.</param>
#<param name="$serviceName">Service name.</param>
#<param name="$requiredStartupMode">Startup mode which must be applied to the service.</param>
#<param name="$timeout">Timeout to check whether the service has been started.</param>
function Stop-Service ([string]$server, [string]$serviceName, [int]$timeout) 
{
	try
	{
		$service = Get-Service $serviceName -ComputerName $server
		if ($service -eq $null) 
		{
			Add-Row $server $row.Service $row.Description "Not found"
			Write-Host "    Service" $serviceName "on" $server "not found" -ForegroundColor Red
			$Global:errorCount++
		}
		else
		{
			if ($service.Status -eq "Running") 
			{
				# If the current service has dependent services, recursively stop them
				[string]$dependentServices = $service.DependentServices
				if ($dependentServices.trim().Length > 0) 
				{
					$dependentServiceList = $dependentServices -split "\s+"
					foreach ($dependentService in $dependentServiceList) 
					{
						Stop-Service $serverName $dependentService $Global:defaultTimeout
					}
				}

				# Record the service state and stop the service
				$Global:lineCount++
				$startupMode = Get-StartupMode $serverName $service.Name
				Add-Row $Global:lineCount $serverName $service.Name $service.DisplayName $service.Status $startupMode $timeout
				$service.Stop()
                
				# Set timeout if specified
				if ($timeout -gt 0) 
				{
					do 
					{
						$service = Get-Service $serviceName -ComputerName $server
						$currentStatus = $service.Status
						Write-Host "        Waiting the service " $serviceName "on" $server "to stop. Current status: " $currentStatus -ForegroundColor Yellow
						Start-Sleep -seconds $timeout
					}
					until ($currentStatus -eq "Stopped")
				}

				Write-Host "    Stopped service " $serviceName "on" $server -ForegroundColor White
			}
		}
	}
	catch
	{
		Write-Host "    Exception on stopping service" $serviceName "on" $server -ForegroundColor Red
		Write-Host $_.Exception -ForegroundColor Red
		$Global:errorCount++
	}
}

#**************************************************************************************
# Main script
#**************************************************************************************

# Assignment of constant values
$columnDelimitor = ","
$scriptFolder = Get-ScriptFolder
$serviceListFile = $scriptFolder + "\Services.csv"
$serviceStatusFile = $scriptFolder + "\ServiceStatus.csv"

if (Test-Path $serviceStatusFile)
{
	#File exists; stop the script
	Write-Host "Status file already exists at" $serviceStatusFile ". Script is stopped to avoid overwriting." -ForegroundColor Red
}
else
{
	Add-Row "Line" "Server" "Service" "Description" "Status" "StartupType" "Timeout"   
	[string]$serverName = "-"
	$input = Import-Csv -path $serviceListFile
	# Iterate through service list file
	ForEach($row in $input)
	{
		if ($serverName -ne $row.Server)
		{
			$serverName = $row.Server
		}
		Stop-Service $serverName $row.Service $row.Timeout 
	}
	if ($Global:errorCount -eq 0)
	{
		Write-Host "Successfully stopped all listed services." -ForegroundColor Green
	}
	else
	{
		Write-Host "Stopped all listed services with" $Global:errorCount "errors." -ForegroundColor Red
	}
}
