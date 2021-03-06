param($statsLoc)
$VMWareStatsLoc = $statsLoc
$currentDate = (get-date).ToString("yyyyMMdd")
$runPathNew = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
$rootpath = (get-item $runPathNew ).parent.FullName
$logfileDate = (get-date).ToString("MMddyy")
$Logfile = "$rootpath\Logs\$(gc env:computername)_$logfileDate.log"
$logDate = (get-date).ToString("MM/dd/yy HH:mm:ss")

Function LogWrite
{
   Param ([string]$logstring)
   #make sure path Exist
   Add-content $Logfile -value "$logDate : $logstring"
}

Set-StrictMode -Version 3
Add-Type -Assembly System.IO.Compression.FileSystem
try{
    # get host stats and zip them up
    $excelFiles = Get-ChildItem $VMWareStatsLoc\Excel\Host\*.xls
    $zipName = $VMWareStatsLoc + "\Archive\Host\VMWare_Stats_" + $currentDate + ".zip"
    LogWrite "Added list of excel files to variable"
    
     #force remove zip file if it already exists
        if(Test-Path $zipName){
            Remove-Item $zipName -Force
            LogWrite "Removed zip file that was found with name $zipName"
        }
    
    #this needs to happen after the remove
        $archive = [System.IO.Compression.ZipFile]::Open($zipName, "Create")
    #loop each file in array
        foreach($file in $excelFiles){
            $item = $file | Get-Item
            $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $item.FullName, $item.Name)
            LogWrite "Added $file to Zip Archive"
        }
    
}
finally{
    $archive.Dispose()
    $archive = $null
    LogWrite "Archive in memory disposed"
}
