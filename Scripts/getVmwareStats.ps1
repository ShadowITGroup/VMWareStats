param($statsLoc,$daysRun,$cliUser,$cliPass)
$runPathRoot = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
#need to get the root path since scripts are 1 directory higher
$rootpath = (get-item $runPathRoot ).parent.FullName
& $rootPath\loadConfig.ps1 $rootPath\vmwareStats.config
$vcenter = $appSettings["VCenterServer"]
Connect-VIServer $vcenter -User $cliUser -Password $cliPass
$allvms = @()
$allhosts = @()
$hosts = Get-VMHost
$vms = Get-Vm
$loc = $statsLoc
#$daysRun = 90
#for testing we will use 2 days
$daysRun = $daysRun

$logfileDate = (get-date).ToString("MMddyy")
$Logfile = "$rootpath\Logs\$(gc env:computername)_$logfileDate.log"
$logDate = (get-date).ToString("MM/dd/yy HH:mm:ss")

Function LogWrite
{
   Param ([string]$logstring)
   #make sure path Exist
   Add-content $Logfile -value "$logDate : $logstring"
}

	foreach($vmHost in $hosts){
		#$daysCount = 1
		#$daysMinus = 2
        #echo out which host we are on
        echo "Host : $vmHost.name"
        
        $daysCount = $daysRun
		$daysMinus = [int]$daysRun + 1
		while($daysCount -ge 0){
		  $hoststat = "" | Select HostName, MemMax, MemAvg, MemMin, CPUMax, CPUAvg, CPUMin, Date
		  $hoststat.HostName = $vmHost.name
		  
          #just echoing the day count so i know where it is at in the loop
          echo "$daysMinus : $daysCount"
          
		  $statcpu = Get-Stat -Entity ($vmHost)-start (get-date).AddDays(-$daysMinus) -Finish (Get-Date).AddDays(-$daysCount) -MaxSamples 100 -stat cpu.usage.average
		  $statmem = Get-Stat -Entity ($vmHost)-start (get-date).AddDays(-$daysMinus) -Finish (Get-Date).AddDays(-$daysCount) -MaxSamples 100 -stat mem.usage.average

		  $cpu = $statcpu | Measure-Object -Property value -Average -Maximum -Minimum
		  $mem = $statmem | Measure-Object -Property value -Average -Maximum -Minimum
		  
		  $hoststat.CPUMax = $cpu.Maximum
		  $hoststat.CPUAvg = $cpu.Average
		  $hoststat.CPUMin = $cpu.Minimum
		  $hoststat.MemMax = $mem.Maximum
		  $hoststat.MemAvg = $mem.Average
		  $hoststat.MemMin = $mem.Minimum
		  $hoststat.Date = (get-date).AddDays(-$daysMinus)
		  
          $daysMinus = $daysCount
		  $daysCount = $daysMinus - 1
		  
		  $allhosts += $hoststat
		}
        LogWrite "getVmwareStats.ps1 retreived stats for $vmHost"
	}
    $fileNameHosts = $loc + "\Hosts.csv"
	$allhosts | Select HostName, MemMax, MemAvg, MemMin, CPUMax, CPUAvg, CPUMin, Date | Export-Csv $fileNameHosts -noTypeInformation

	#foreach($vm in $vms){
	#	$daysCount = 1
	#	$daysMinus = 2
	#	while($daysCount -le $daysRun){
	#	  $vmstat = "" | Select VmName, MemMax, MemAvg, MemMin, CPUMax, CPUAvg, CPUMin, Date
	#	  $vmstat.VmName = $vm.name
	#	  
	#	  $statcpu = Get-Stat -Entity ($vm)-start (get-date).AddDays(-$daysMinus) -Finish (Get-Date).AddDays(-$daysCount) -MaxSamples 100 -stat cpu.usage.average
	#	  $statmem = Get-Stat -Entity ($vm)-start (get-date).AddDays(-$daysMinus) -Finish (Get-Date).AddDays(-$daysCount) -MaxSamples 100 -stat mem.usage.average
	#
	#	  $cpu = $statcpu | Measure-Object -Property value -Average -Maximum -Minimum
	#	  $mem = $statmem | Measure-Object -Property value -Average -Maximum -Minimum
	#	  
	#	  $vmstat.CPUMax = $cpu.Maximum
	#	  $vmstat.CPUAvg = $cpu.Average
	#	  $vmstat.CPUMin = $cpu.Minimum
	#	  $vmstat.MemMax = $mem.Maximum
	#	  $vmstat.MemAvg = $mem.Average
	#	  $vmstat.MemMin = $mem.Minimum
	#	  $vmstat.Date = (get-date).AddDays(-$daysMinus)
	#	  
	#	  $daysCount = $daysCount + 1
	#	  $daysMinus = $daysCount + 1
	#	  $allvms += $vmstat
	#	}
	#}
	#$allvms | Select VmName, MemMax, MemAvg, MemMin, CPUMax, CPUAvg, CPUMin, Date | Export-Csv "c:\VMs.csv" -noTypeInformation