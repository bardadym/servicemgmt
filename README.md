# servicemgmt

Manage services on servers by PowerShell scripts

Series of three scripts are proposed to manage services on specified servers in a farm. 
Windows and application services may require to be stopped for the time of migration, installation of updates, or other operations. 
The scripts are using CSV files for input and output. All input and output CSV files are supposed to be located in the same folders as scripts. 
 
The script List-Services.ps1 reads the list of servers from the input file Servers.csv and writes the list of Windows services on all listed servers to the file ServicesFullList.csv. 
This script has the following purpose: 
•	To identify application services and to build the list of services to stop / start by other scripts. 
•	To make a list of services with their statuses before their bulk stopping. 
•	To control the status of services after their stopping and restarting by other scripts. 

CSV file ServicesFullList.csv has the following information: 
1.	Server, 
2.	Service name, 
3.	Service description, 
4.	Status of the service, 
5.	Startup type of the service, 
6.	Dependent services (space-delimited service names),
7.	Required services (space-delimited service names),
8.	Timeout (foreseen to specify timeout in seconds needed to start some services). 
You can copy this file to Services.csv, and delete rows for the services which you don’t want to stop / start. Some services require a timeout to stop / start. This timeout may be specified in seconds in the column Timeout in the input file Services.csv.

The script Stop-Services.ps1 reads the list of services to stop from the input file Services.csv, stops running services and writes the list of stopped servers to the file ServiceStatus.csv. 
If the output file ServiceStatus.csv already exists, the script stops working, supposing that an earlier run of this script has stopped the listed services. 
If a service to stop has dependencies, other services which require it running, these services (and their dependencies) are being stopped by a recursive call of the script. 

The script Start-Services.ps1 reads the list of stopped services from the input file ServiceStatus.csv, and starts the listed services in reversed order, in other words, "Last stopped - first started".
This way, the script keeps natural dependencies between servers and services. For example, services on web and application servers must be stopped before and started after the services on backend (database servers).
After starting all listed services, the script deletes the file ServiceStatus.csv. 
If a service to start has dependencies, other services required to run, these services (and their dependencies) are being started by a recursive call of the script. 

Scenario of use.

All scripts must run with elevated permissions (as administrator). The account which is running the script must have administrator permissions to all listed servers. 
1.	Prepare the list of servers in the file Servers.csv. Keep the order of the servers following the dependencies of services across them (e.g. web front end - application - database).  
2.	Run the script List-Services.ps1. 
3.	Make a copy of the file ServicesFullList.csv to ServicesFullListBefore.csv. 
4.	Rename the file ServicesFullList.csv to Services.csv. 
5.	Edit the file Services.csv. Leave only the rows of the services to be stopped / started (for example, leave only SharePoint and SQL Server services for SharePoint farm). Delete critical Windows services from the list, as their stopping can bring down the servers. (see the next chapter)
6.	Run the script Stop-Services.ps1.
7.	Perform required maintenance actions (upgrade, installation, migration,...).
8.	Run the script Start-Services.ps1.
9.	Assure that applications are working. 
10.	Run the script List-Services.ps1.
11.	Rename the file ServicesFullList.csv to ServicesFullListAfter.csv. 
12.	Compare the files ServicesFullListBefore.csv and ServicesFullListAfter.csv in Excel and assure that all services have been started in the right mode, as they were before stopping. 

Attention: Do not stop these services.

Power is always dangerous. Assure that the input of the scripts does not include the following Windows services, which are critical for functioning of Windows and connections to the server. 
1.	Appinfo (Application Information), 
2.	LanmanServer and all its dependencies (both required and dependent services), 
3.	LanmanWorkstation and all its dependencies (both required and dependent services), 
4.	Netlogon, 
5.	Netman (Network Connections), 
6.	NlaSvc (Network Location Awareness), 
7.	PlugPlay (Plug and Play), 
8.	SamSs (Security Accounts Manager), 
9.	Seclogon (Secondary Logon), 
10.	TermService (Remote Desktop Services) and all its dependencies (both required and dependent services), 
11.	W32Time (Windows Time). 

The list is not comprehensive. 
