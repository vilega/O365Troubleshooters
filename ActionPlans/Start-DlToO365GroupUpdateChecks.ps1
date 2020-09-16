<# 1st requirement install the module O365 TS
Import-Module C:\Users\a-haemb\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>
Clear-Host
#region Connecting to EXO & MSOL

$Workloads = "exo","msol"
try {
    Connect-O365PS $Workloads 
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
catch {
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Failure"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
#endregion Connecting to EXO & MSOL

    # Create working folder for Groups Diag
    $ts= get-date -Format yyyyMMdd_HHmmss
    $ExportPath = "$global:WSPath\_$ts"
    mkdir $ExportPath -Force |out-null

    <# Conditions #>

# End of the Diag
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu
