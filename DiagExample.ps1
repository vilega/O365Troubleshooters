# diag example
Write-Host "Test"
Import-Module 'C:\Users\vilega\Documents\GitHub\O365Troubleshooters\CommonFunctions.psm1' -DisableNameChecking
Import-Module 'C:\Users\vilega\Desktop\O365 troubleshooters\CommonFunctions.psm1' -DisableNameChecking
Import-Module 'C:\Users\vilega\Documents\GitHub\O365Troubleshooters\CommonFunctions.psm1' -DisableNameChecking
Write-Host "After module import"
#Get-Command -Module CommonFunctions.psm1
#Add-ScriptBlocks
Request-Credential
