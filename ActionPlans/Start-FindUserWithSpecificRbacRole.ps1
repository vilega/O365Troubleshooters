Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RbacRoleSpecific_$ts"
mkdir $ExportPath -Force | Out-Null
. $script:modulePath\ActionPlans\Start-RbacTools.ps1
Get-SpecificRoleMembers |export-csv "$ExportPath\RoleMembers_$ts.csv" -NoTypeInformation 
Write-Host "The list of user who have selected roles assigned was exported to $ExportPath."
Read-Key
Start-O365TroubleshootersMenu