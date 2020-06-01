Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RbacRoleAll_$ts"
mkdir $ExportPath -Force |Out-Null
. $script:modulePath\ActionPlans\Start-RbacTools.ps1
Get-AllUsersWithAllRoles | export-csv "$ExportPath\ManagementRoleAssignmentUsers_$ts.csv" -NoTypeInformation
Write-Host "Export all users with all the roles assigned to the file: $ExportPath"
Read-Key
Start-O365TroubleshootersMenu