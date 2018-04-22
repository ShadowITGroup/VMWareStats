Before Running RunStats.EXE.ps1 please read this.
Author: Matthew Davis
----------------------------------------------------------------------

REQUIRED Installs!!!!!

	1. PsTools needed on remote server executing powershell scripts.
	2. MSChart.exe needs to be installed on server that has excel installed on it.
	3. powershell 3.0 is required on both remote server and local
	4. .NET 4.5 framework required on server running scripts

!!!!!!!!**********************************************!!!!!!!!!!
	Modify vmwareStats.config to fit your environment
!!!!!!!!**********************************************!!!!!!!!!!

-------------------------------------------------------------------------------------

To run vmware Stats program :
--------------------------------------------------------------------------------------
	Run RunStats.EXE.ps1 in scheduler or powershell. 

	This program will utilize the other scripts written to
	compile stats from VMWare, graph data that is retreived, place data and graph for each node in file, zip it up, 
	and email it to email that is defined in vmWareStats.config
---------------------------------------------------------------------------------------

EX:
--------------------------------------------------------------------------------
	Server A will run the PowerShell 3.0 and VSphere PowerCLI program but does not have Excel installed on it. So it needs 
to call server B to use excel to chart Data and put it into an excel spreadsheet for us. 

We will extract PSTools into C:\PSTools on Server B and install Powershell 3.0 and MSChart on Server B.Then we will install Powershell 3.0 
and VSphere PowerCLI on Server A. 


Description:
---------------------------------------------------------------------------------
Server A will do all the processing of the data, and will only use Server B for its Excel engine. Server A uses PSTools to remotely
execute commands on Server B in the background. Once done, Server A will Zip the files that were created in the VmwareStatsLoc, 
and email them to the email defined in the config file.


	VCenterServer
	---------------------------------------------
		This variable is the name of the VCenter server that this program will connect to.

	RunStats
	---------------------------------------------
		When set to 0, the initial part of the program that grabs data from VMWare will be skipped. This is good to use
		If you don't want to re run the stats from vmware, but want it to go through re creating the spreadsheets, Zip archiving,
		Then email the results again. Else if set to 1 it will connect to VCenter and pull stats.

	PSToolsDir
	----------------------------------------------
		This is the directory in which PSTools is installed on the remote machine. By default it is set to C:\PSTools because that
		is where i usually extract pstools to.

	RemoteRunXL
	----------------------------------------------
		Since this process can be cpu intensive. It is recommended that you run this on a non production server. However this program
		requires Excel in order to put data into a spreadsheet. With this value set to 1, you can run this program on 1 server, that
		will do all the heavy work and then run scripts remotely on a server that has Excel installed on it to ease the cpu time. 
		I use this on a non production server, and have it call to a Xen-App server that has excel installed to create the spreadsheets.
		If set to 0 it will run everything locally.

	XLServerScriptUNC
	----------------------------------------------
		This is the UNC path of where only the necessary files needed to create the spreadsheets will be copied to. It should be in
		UNC form. EX: \\ServerName\c$\VMWare\ NOTE: trailing slash.

	XLLocalScriptPath
	----------------------------------------------
		This value should be the same location as above, but instead of providing the UNC path you will give the local path. EX. C:\VMWare

	XLServerName
	----------------------------------------------
		If RunRemoteXL is set to 1 then this is the server name that you would like the script to use for executing Excel scripts.

	enableAddZipExt
	----------------------------------------------
		Some mail clients remove email attachments that have certain extensions. This enables you to tell the script to change the file
		extension of the attachment so that the zip fil will not be removed.When set to 1 the script will assign a new extension to the file
		before it emails it.

	addZipExt
	----------------------------------------------
		When enableAddZipExt is set to 1 the zip file extension will be changed to whatever the value of this variable is. If the value is
		"TXT" then the file will be named filename.zip.txt

	PSToolsUser
	-----------------------------------------------
		This is the user that will be used to execute PSExec on the remote server. This user should have network access, and administrative 
		access on the remote server to execute PSExec scripts.

	PSToolsPass
	-----------------------------------------------
		The value of this is the password for PSToolsUser. This is cleartext unfortunately. If i have more time i will try to determine a way
		to obfustcate it.

	VmwareStatsLoc
	-----------------------------------------------
		This is the location where the VMWareStats and archives will be stored. It can be a local or remote unce path.

	DaysRun
	-----------------------------------------------
		Self explanantory, but is the number of days that you would like the scripts to capture. If you want the program to get stats for 
		the last 3 months for each VMWare host, then set this value to 90.

	CLIUser
	-----------------------------------------------
		This is the user that has administrative access to the vcenter server. It will be used in VMWare PowerCLI to connect to VCenter
		database.

	CLIPass
	-----------------------------------------------
		The password to be used with CLIUser. I am trying to determine how to obfuscate the password in the config so that it does 
		not sit in the config in cleartext form.

	EmailEnable
	-----------------------------------------------
		This program has the ability to send an email out after the process has been completed and the zip file archive has been created.
		Set this to 1 if you want the program to email you after it is complete.

	SmtpServer
	-----------------------------------------------
		This is the SMTP server to use to send the email.

	EmailFrom
	-----------------------------------------------
		The email that is sent will have this address in the from field.

	EmailTo
	-----------------------------------------------
		Email adresses to send the attachment to. This can be a single value, or a comma delimited list.

	EmailSubject
	-----------------------------------------------
		You can customize the subject line of the email using this value.
		