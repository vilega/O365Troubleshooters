$Workloads = "exo"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RbacRole_$ts"
mkdir $ExportPath -Force
. $script:modulePath\ActionPlans\Start-RbacTools.ps1
Get-AllUsersWithAllRoles
Read-Host "Press any key then [Enter] to return to main menu"
Clear-Host
Start-O365TroubleshootersMenu