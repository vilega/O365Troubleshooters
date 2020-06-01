$Workloads = "exo"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$global:ExportPath = "$global:WSPath\RbacRole_$ts"
mkdir $ExportPath -Force |Out-Null
. $script:modulePath\ActionPlans\Start-RbacTools.ps1
Get-AllUsersWithAllRoles
Read-Key
Start-O365TroubleshootersMenu