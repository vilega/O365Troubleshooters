Import-Module "C:\Users\vilega\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1" -Force
#start-O365Troubleshooters
Set-GlobalVariables
Connect-O365PS -O365Service AipService
Connect-O365PS -O365Service Scc
#Connect-O365PS -O365Service Exo
#Connect-O365PS -O365Service Msol
