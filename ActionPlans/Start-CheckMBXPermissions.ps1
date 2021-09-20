
<# 
Import-Module C:\Users\alexaca\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
Start-O365TroubleshootersMenu
#>
# connect
Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads

# logging
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\MailboxDiagnosticLogs_$ts"
mkdir $ExportPath -Force |out-null


$mbx = Read-Host "SMTP "
write-host (Get-Mailbox $mbx).Name



Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu