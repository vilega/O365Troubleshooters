# diag example
Write-Host "Test"
Import-Module 'C:\Users\vilega\Documents\GitHub\O365Troubleshooters\CommonFunctions.psm1' -DisableNameChecking

Import-Module CommonFunctions
CommonFunctions\MENU

Write-Host "After module import"
#Get-Command -Module CommonFunctions.psm1
#Add-ScriptBlocks
Request-Credential
Connect-O365PS exo
