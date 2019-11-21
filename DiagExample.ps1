# diag example
Write-Host "Test"
Import-Module .\CommonFunctions.psm1 -DisableNameChecking
Write-Host "After module import"
#Get-Command -Module CommonFunctions.psm1
#Add-ScriptBlocks
Request-Credential
