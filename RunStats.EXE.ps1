#############################
# Depedencies: MSChart.exe is required on the server that will be running graphVmwareStats.ps1
#              The Server that runs compileSpreadSheet.ps1 needs to have excel installed on it.
###
# copy scripts to location so that we can run them against excel since these are xenapp servers, the files will be deleted on reboot, so we need to copy them every time we want to run them.
# We need to do this because xenapp servers are the only servers that have Excel on them to process the data.
#
# load config into memory
## Adds one or more Windows PowerShell snap-ins to the current session.
Add-PSSnapin VMware.VimAutomation.Core
Add-PSSnapin VMware.VimAutomation.License
Add-PSSnapin VMware.DeployAutomation
Add-PSSnapin VMware.ImageBuilder 
## Retrieves the versions of the installed PowerCLI snapins.
Get-PowerCLIVersion

#Use this for Logging
$runPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
#load config
. $runPath\loadConfig.ps1 $runPath\vmwareStats.config

$VMWareStatsLoc = $appSettings["VmwareStatsLoc"]
$logfileDate = (get-date).ToString("MMddyy")
$Logfile = "$runPath\Logs\$(gc env:computername)_$logfileDate.log"
$PSLogfile = "$runPath\Logs\PSTools_$(gc env:computername)_$logfileDate.log"
$excelServer = $appSettings["XLServerScriptUNC"]
$localPath = $appSettings["XLLocalScriptPath"]
$daysRun = $appSettings["DaysRun"]
$runStats = $appSettings["RunStats"]
$cliUser = $appSettings["CLIUser"]
$cliPass = $appSettings["CLIPass"]
$runRemoteXL = $appSettings["RunRemoteXL"]
$psToolsDir = $appSettings["PSToolsDir"]
$remoteServer = "\\" + $appSettings["XLServerName"]
$currentDate = (get-date).ToString("MM/dd/yy HH:mm:ss")

# include test ps script
. "$runPath\test.ps1"
$checkMSChart = checkMSChartAssembly
$checkZip = checkZipAssembly

if($checkMSChart -and $checkZip){
        Function LogWrite
        {
           Param ([string]$logstring)
           #make sure path Exist
           if (-Not(Test-Path  ("$runPath\Logs"))){ New-Item -ItemType directory -Path "$runPath\Logs"}
           Add-content $Logfile -value "$currentDate : $logstring"
        }
    
    #make sure directories exist
    if (-Not(Test-Path  ($VMWareStatsLoc))){ New-Item -ItemType directory -Path $VMWareStatsLoc}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Archive"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Archive"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Archive\Host"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Archive\Host"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Archive\VM"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Archive\VM"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Excel"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Excel"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Excel\Host"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Excel\Host"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Excel\VM"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Excel\VM"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Images"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Images"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Images\Host"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Images\Host"}
    if (-Not(Test-Path  ("$VMWareStatsLoc\Images\VM"))){ New-Item -ItemType directory -Path "$VMWareStatsLoc\Images\VM"}
    if (-Not(Test-Path  ("$excelServer"))){ New-Item -ItemType directory -Path "$excelServer"}

    #run Initial script to compile data into csv file
    if($runStats -eq 1){
        . $runPath\Scripts\getVmwareStats.ps1 -statsLoc $VMWareStatsLoc -daysRun $daysRun -cliUser $cliUser -cliPass $cliPass
    }
    copy-item -Path $runPath\Installs\regKey.reg -Destination $excelServer
    LogWrite "regKey.reg copied to destination"

    copy-item -Path $runPath\Scripts\compileSpreadSheet.ps1 -Destination $excelServer
    LogWrite "compileSpreadSheet.ps1 copied to destination"

    copy-item -Path $runPath\Scripts\graphVmwareStats.ps1 -Destination $excelServer
    LogWrite "graphVmwareStats.ps1 copied to destination"

    #need to delete image files so new ones can be created. Also delete csv host files to recreate.
        $imageFiles = Get-ChildItem $VMWareStatsLoc\Images\Host\*.png
        foreach($file in $imageFiles){
    	    Remove-Item $file
            LogWrite "$file has been removed"
        }

        $excelFiles = Get-ChildItem $VMWareStatsLoc\Excel\Host\*.xls
        foreach($file in $excelFiles){
    	    Remove-Item $file
            LogWrite "$file has been removed"
        }

    #modify reg key
    if($runRemoteXL -eq 1){
        LogWrite "Running Command : regedit /s $runPath\Installs\regKey.reg on $remoteServer"
        . "$psToolsDir\psexec" $remoteServer -u $appSettings["PSToolsUser"] -p $appSettings["PSToolsPass"] /accepteula regedit /s "$localPath\regKey.reg" 2>>$PSLogfile
        LogWrite "Running Command : powershell.exe -file $localPath\graphVmwareStats.ps1 -statsLoc $VMWareStatsLoc\ on $remoteServer"
        . "$psToolsDir\psexec" $remoteServer -u $appSettings["PSToolsUser"] -p $appSettings["PSToolsPass"] /accepteula cmd.exe /c "echo . | powershell.exe -file $localPath\graphVmwareStats.ps1 -statsLoc $VMWareStatsLoc\" 2>>$PSLogfile
        #now that the images have been made we will run a script that creates the excel file and includes the image in it.
        LogWrite "Running Command : powershell.exe -file $localPath\compileSpreadSheet.ps1 -statsLoc $VMWareStatsLoc\ on $remoteServer"
        . "$psToolsDir\psexec" $remoteServer -u $appSettings["PSToolsUser"] -p $appSettings["PSToolsPass"] -i 0 /accepteula cmd.exe /c "echo . | powershell.exe -file $localPath\compileSpreadSheet.ps1 -statsLoc $VMWareStatsLoc\" 2>>$PSLogfile
    }else{
        LogWrite "Running graphVmwareStats.ps1"
        . $runPath\Scripts\graphVmwareStats.ps1 -statsLoc "$VMWareStatsLoc\"
        LogWrite "Running compileSpreadSheet.ps1"
        . $runPath\Scripts\compileSpreadSheet.ps1 -statsLoc "$VMWareStatsLoc\"
    }

    #now we will run a script that will create the zip archive that can be used to email stats.
    LogWrite "Running Command : \Scripts\zipVmwareStats.ps1 -statsLoc $VMWareStatsLoc"
    . $runPath\Scripts\zipVmwareStats.ps1 -statsLoc $VMWareStatsLoc
    #now that it is zipped we will email it to email listed in config.
    LogWrite "Running Command : \Scripts\emailZip.ps1 -statsLoc $VMWareStatsLoc"
    . $runPath\Scripts\emailZip.ps1 -statsLoc $VMWareStatsLoc
}else{
    if(-Not($checkMSChart)){
        Throw "Please make sure you have MSChart installed on server running excel"
    }
    if(-Not($checkZip)){
        Throw "Missing Assembly!. Please make sure you have .NET 4.5 CLR installed locally"
    }
}