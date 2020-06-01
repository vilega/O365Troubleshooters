Import-Module "C:\Users\vilega\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1" -Force
start-O365Troubleshooters
Set-GlobalVariables
Connect-O365PS -O365Service Exo
Connect-O365PS -O365Service Msol