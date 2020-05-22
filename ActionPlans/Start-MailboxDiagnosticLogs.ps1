
    
# Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
$Workloads = "exo"
Connect-O365PS $Workloads
    
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
# Main Function
$Ts= get-date -Format yyyyMMdd_HHmmss

Write-Host "`nPlease input the path were the files will be saved" -ForegroundColor Green
$ExportPath = Read-Host

if ($ExportPath[-1] -eq "\") {
    $ExportPath = $ExportPath.Substring(0,$ExportPath.Length-1)
}

If (Test-Path -Path $ExportPath) {
    #Write-Host "`nThe path exist!" -ForegroundColor Green
}
else {
    Write-Host "`nThe output folder doesn't exist or is not valid! Please create or use an existing one and re-run the script. Press [Enter] to exit" -ForegroundColor Red
    Read-Host
    Exit
}

#endregion

#region MbxDiagLogs

Write-Host "`nPlease input the mailbox for which you want to see MailboxDiagnosticLogs: " -ForegroundColor Green
$mbx = Read-Host

# Check if mailbox exist
$previousErrorActionPreference = $global:ErrorActionPreference
$global:ErrorActionPreference = 'Stop'
try{
    Get-Mailbox $mbx | Out-Null
}
Catch{
    Write-Host "`nThe mailbox $mbx doesn't exist. Press [Enter] to exit"
    Read-Host
    $ErrorActionPreference = $previousErrorActionPreference
    Exit
}
$global:ErrorActionPreference = $previousErrorActionPreference


# Getting available components that can be exported 
$previousErrorActionPreference = $global:ErrorActionPreference
$global:ErrorActionPreference = 'Stop'
$global:error.Clear()
Try {
    Export-MailboxDiagnosticLogs $mbx -ComponentName TEST 
}
Catch {
    #Write-Host "in catch"
    $global:MbxDiagLogs = ((($global:error[0].Exception.Message -Split "Available logs: ")[1] -replace "'") -split ",") -replace " "
}

$global:ErrorActionPreference = $previousErrorActionPreference

    # Export-MailboxDiagnosticLogs with ComponentName
$option = ( $global:MbxDiagLogs + "ALL")|Out-GridView -PassThru -Title "Choose a specific ComponentName or the last one for ALL"
if ($option -ne "ALL") {
    Write-Host "`nGetting $option logs" -ForegroundColor Yellow 
    $option | ForEach-Object {
        Export-MailboxDiagnosticLogs $mbx -ComponentName  $_ | Tee-Object $ExportPath\$($Ts)_$_.txt
    } 
}
else {
    $MbxDiagLogs |ForEach-Object{
        Write-Host "`nGetting $_ logs" -ForegroundColor Yellow 
        Export-MailboxDiagnosticLogs $mbx -ComponentName  $_ | Tee-Object $ExportPath\$($Ts)_$_.txt
    }
}


# Export-MailboxDiagnosticLogs with ExtendedProperties
Write-Host "You can view & filter ExtendedProperties in the Grid View window." -ForegroundColor Yellow
$extendLogs = Export-MailboxDiagnosticLogs $mbx -ExtendedProperties
$ExtendedProps = [XML]$extendLogs.MailboxLog
$ExtendedProps.Properties.MailboxTable.Property | Select-Object name,value | Out-GridView -Title "All ExtendedProperties with values (you can filter here to find what is interesting for you; e.g: use `"ELC`" for MRM properties)"
$ExtendedProps.Properties.MailboxTable.Property | Select-Object name,value | Out-File $ExportPath\$($Ts)_ExtendedProperties.txt

Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 

# Disconnecting
Disconnect-all  

# Go back to the main menu
Start-O365TroubleshootersMenu
