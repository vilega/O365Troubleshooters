# diag example
Write-Host "Test"
Import-Module C:\Users\vilega\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -DisableNameChecking
Import-Module C:\Users\vilega\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Verbose -Force

Import-Module CommonFunctions.psm1 -Force
CommonFunctions\MENU

Write-Host "After module import"
#Get-Command -Module CommonFunctions.psm1
#Add-ScriptBlocks
Request-Credential
Connect-O365PS exo


Publish-Module -Name "O365Troubleshooters" -Path "C:\Users\vilega\Documents\GitHub\O365Troubleshooters\"  -NuGetApiKey "oy2h3ibtm6jsynifz5mqmm6at5gkownefjuz7urc7qwlgm" -ProjectUri "https://github.com/VictorLegat/O365Troubleshooters"

Publish-Module -Name "O365Troubleshooters"  -NuGetApiKey "oy2h3ibtm6jsynifz5mqmm6at5gkownefjuz7urc7qwlgm" -ProjectUri "https://github.com/VictorLegat/O365Troubleshooters"

New-ModuleManifest -Path "C:\Program Files\WindowsPowerShell\Modules\O365Troubleshooters\O365Troubleshooters.psd1"
Update-ModuleManifest -ProjectUri "https://github.com/VictorLegat/O365Troubleshooters" -Path "C:\Program Files\WindowsPowerShell\Modules\O365Troubleshooters\O365Troubleshooters.psd1"
Register-PSRepository -Name "O365Troubleshooters" -SourceLocation "https://github.com/VictorLegat/O365Troubleshooters" -PublishLocation "https://github.com/VictorLegat/O365Troubleshooters" -InstallationPolicy Trusted
