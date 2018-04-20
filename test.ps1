$runPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
. $runPath\loadConfig.ps1 $runPath\vmwareStats.config
$psToolsDir = $appSettings["PSToolsDir"]
$remoteServer = "\\" + $appSettings["XLServerName"]
$excelServer = $appSettings["XLServerScriptUNC"]
$localPath = $appSettings["XLLocalScriptPath"]

Function checkZipAssembly
{
    Param([bool]$found)

    Try 
    {   # we use the -ea to tell the cmdlet to silence any non-terminating errors since try catch only handles terminating errors.
        Add-Type -AssemblyName System.IO.Compression.Filesystem -ea SilentlyContinue
        return $true
    }
    Catch [Exception]
    { 
        return $false
    }
    #Finally {    
    #    $clrVer = [System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion        
    #    if($found -eq $false){
    #        Throw "Please make sure you have .NET 4.5 CLR installed!"
    #    }
    #}
}
Function checkMSChartAssembly
{
    ####!!!!Dont use this!!! ##
    # For 32bit look for key
    # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41785C66-90F2-40CE-8CB5-1C94BFC97280}
    # For 64 bit
    # HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432node\Microsoft\Windows\CurrentVersion\Uninstall\{41785C66-90F2-40CE-8CB5-1C94BFC97280}
    #
    #get registry key for Microsoft Chart Controls
    Try{
        if (-Not(Test-Path  ("$excelServer"))){ New-Item -ItemType directory -Path "$excelServer"}
        copy-item -Path $runPath\Scripts\checkMSChart.ps1 -Destination $excelServer
        . "$psToolsDir\psexec" $remoteServer -u $appSettings["PSToolsUser"] -p $appSettings["PSToolsPass"] /accepteula cmd.exe /c "echo . | powershell.exe -file $localPath\checkMSChart.ps1" > "$runPath\remoteInstalls.txt" 2> "$runPath\Logs\MSChartTest.txt"
        
        [string]$parseIt = Get-Content $runPath\remoteInstalls.txt | Select-String -Pattern 'Microsoft Chart'
        
        if($parseIt.Length -gt 0){
            return $true
        }else{
            return $false
        }
    }
    Catch [Exception]{
        return $false
    }
}
