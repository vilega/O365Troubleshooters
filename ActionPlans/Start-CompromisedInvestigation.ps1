Function Start-CompromisedMain
{


}
$Workloads = "exo","SCC", "MSOL"#, "AAD"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\Compromised_$ts"
mkdir $ExportPath -Force

. $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1


Start-CompromisedMain