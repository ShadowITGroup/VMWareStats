$runPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
#need to get the root path since scripts are 1 directory higher
$rootpath = (get-item $runPath ).parent.FullName

& $rootPath\loadConfig.ps1 $rootPath\vmwareStats.config
$emailEnable = $appSettings["EmailEnable"]
$statsLoc = $appSettings["VmwareStatsLoc"]
$smtpServer = $appSettings["SmtpServer"]
$emailFrom = $appSettings["EmailFrom"]
$emailTo = $appSettings["EmailTo"]
$enableRename = $appSettings["enableAddZipExt"]
$addZipExt = $appSettings["addZipExt"]
$emailSubject = $appSettings["EmailSubject"]
$currentDate = (get-date).ToString("yyyyMMdd")
$file = $statsLoc + "\Archive\Host\VMWare_Stats_" + $currentDate + ".zip"
if($enableRename -eq 1){
    $newName = "VMWare_Stats_" + $currentDate + ".zip.$addZipExt"
}else{
    $newName = "VMWare_Stats_" + $currentDate + ".zip"
}
$logfileDate = (get-date).ToString("MMddyy")
$Logfile = "$rootpath\Logs\$(gc env:computername)_$logfileDate.log"
$logDate = (get-date).ToString("MM/dd/yy HH:mm:ss")

Function LogWrite
{
   Param ([string]$logstring)
   #make sure path Exist
   Add-content $Logfile -value "$logDate : $logstring"
}

if($emailEnable -eq 1){
    if($enableRename -eq 1){ Rename-Item $file $newName }
    #after we rename it we will attach it.
    if($enableRename -eq 1){
        $file = $statsLoc + "\Archive\Host\VMWare_Stats_" + $currentDate + ".zip.$addZipExt"
    }else{
        $file = $statsLoc + "\Archive\Host\VMWare_Stats_" + $currentDate + ".zip"
    }
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $att = new-object Net.Mail.Attachment($file)
    $msg.From = $emailFrom
    $msg.Subject = $emailSubject

    foreach($email in $emailTo){
        $msg.To.Add($email)
    }

    $msg.Body = "Attached is the VMWare Stats Requested. Feel free to review at your liesure. Rename and remove .txt on the attachment to view files."

    if (Test-Path  ($file))
    { 
        $msg.Attachments.Add($att)
        $smtp.Send($msg)
        $att.Dispose()
        #now we rename it back so that the archive will show a zip file instead of a txt file.

        $newName = "VMWare_Stats_" + $currentDate + ".zip"
        if($enableRename -eq 1){ Rename-Item $file $newName }
        LogWrite "Email Sent to $emailTo"
    } else {
        LogWrite "$file not found : Cannot attach, so i am not emailing anything"
    }
}else{
    LogWrite "EmailEnable in config set to $emailEnable. No email sent"
}