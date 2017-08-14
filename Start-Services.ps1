# Start Windows services listed in the input file ServiceStatus.csv
# 
# 2017 Viktor Bardadym

#**************************************************************************************
# Constants
#**************************************************************************************

#**************************************************************************************
# Global variables
#**************************************************************************************

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

# Get service startup mode
function Get-StartupMode([string]$server, [string]$serviceName) 
{
	$serviceNameFilter = "Name='" + $serviceName + "'"
	$serviceWmi = Get-WmiObject -Class Win32_Service -Property StartMode -Filter $serviceNameFilter -ComputerName $server
	$serviceWmi.StartMode
}

#<summary>
# Recursively starts services required by the service, then starts the service
#</summary>
#<param name="$server">Server name.</param>
#<param name="$serviceName">Service name.</param>
#<param name="$requiredStartupMode">Startup mode which must be applied to the service.</param>
#<param name="$timeout">Timeout to check whether the service has been started.</param>
function Start-Service ([string]$server, [string]$serviceName, [string]$requiredStartupMode, [int]$timeout) 
{
	try
	{
		$service = Get-Service $serviceName -ComputerName $server
		if ($service -eq $null) 
		{
			Write-Host "    Service" $serviceName "on" $server "not found" -ForegroundColor Red
			$Global:errorCount++
		}
		else
		{
			if ($service.Status -ne "Running") 
			{
				# If the current service has other services required, recursively start them
				[string]$requiredServices = $service.RequiredServices
				if ($requiredServices.trim().Length > 0) 
				{
					$requiredServiceList = $requiredServices -split "\s+"
					foreach ($requiredService in $requiredServiceList) 
					{
						Start-Service $serverName $requiredService $Global:defaultTimeout
					}
				}

				# Start the current service
				$service.Start()

				# Check startup mode with saved before the stop and adjust if needed
				$actualStartupMode = Get-StartupMode $serverName $service.Name
				if ($actualStartupMode -ne $requiredStartupMode) 
				{
					Set-Service -Name $serviceName -StartupType $requiredStartupMode
				}

				# Set timeout if specified
				if ($timeout -gt 0) 
				{
					do 
					{
						$service = Get-Service $serviceName -ComputerName $server
						$currentStatus = $service.Status
						Write-Host "        Waiting the service " $serviceName "on" $server "to start. Current status: " $currentStatus -ForegroundColor Yellow
						Start-Sleep -seconds $timeout
					}
					until ($service.Status -eq "Running")
				}

				Write-Host "    Started service " $serviceName "on" $server "." -ForegroundColor Green
			}
			else
			{
				Write-Host "    Listed service " $serviceName "on" $server "is already running." -ForegroundColor Yellow
			}
		}
	}
	catch
	{
		Write-Host "    Exception on starting service" $serviceName "on" $server -ForegroundColor Red
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
$serviceStatusFile = $scriptFolder + "\ServiceStatus.csv"

if (Test-Path $serviceStatusFile)
{
	# Iterate through service status list file in reverse order
	$input = Import-Csv -path $serviceStatusFile | Sort Line -Descending
	ForEach($row in $input)
	{
		Start-Service $row.Server $row.Service $row.StartupType $row.Timeout
	}
    
	if ($Global:errorCount -eq 0)
	{
		Write-Host "Successfully started all listed services." -ForegroundColor Green
		# Delete status file upon completion
		Remove-Item -path $serviceStatusFile -Force
	}
	else
	{
		Write-Host "Started all listed services with" $Global:errorCount "errors. Please control the script output." -ForegroundColor Red
	}
}
else
{
	# File does not exist; stop the script
	Write-Host "Status file does not exist at" $serviceStatusFile -ForegroundColor Red
}
