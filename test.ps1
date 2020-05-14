Import-Module "C:\Users\vilega\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1" -Force
Import-Module "C:\Users\alnita\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1" -Force

Set-GlobalVariables

$error.Clear()

$InitialErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Stop"
$PSScriptRoot
$ErrorActionPreference = $InitialErrorActionPreference

$CurrentProperty = "Connecting to EXO"
$CurrentDescription = "Success"

#Example Writing Log
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

# Example writting progress on screen
write-progress -activity "Script in Progress" -status "30% Complete: Configuring Global Variables" -percentcomplete 30
Start-Sleep -Milliseconds 500

# Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
$Workloads = "exo","sco","aadrm"
Connect-O365PS $Workloads

# Executing action plan

# Sending collected information
Send-CollectedInfo

# Disconnecting
disconnect-all  