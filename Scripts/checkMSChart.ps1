$a = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | where-object { $_.PSChildName -eq "{41785C66-90F2-40CE-8CB5-1C94BFC97280}"}
$a.DisplayName
if([string]::IsNullOrEmpty($a.DisplayName)){
    $a = Get-ChildItem HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | where-object { $_.PSChildName -eq "{41785C66-90F2-40CE-8CB5-1C94BFC97280}"}
    $a.DisplayName
}