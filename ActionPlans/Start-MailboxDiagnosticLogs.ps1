
    
# Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
$Workloads = "exo"
Connect-O365PS $Workloads
    
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\MailboxDiagnosticLogs_$ts"
mkdir $ExportPath -Force

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
# $global:error.Clear()
$myerror = $null
Try {
    Export-MailboxDiagnosticLogs $mbx -ComponentName TEST 
}
Catch {
    $myerror = $_
    #Write-Host "in catch"
    #$global:MbxDiagLogs = ((($global:error[0].Exception.Message -Split "Available logs: ")[1] -replace "'") -split ",") -replace " "
    $global:MbxDiagLogs = ((($myerror.Exception.Message -Split "Available logs: ")[1] -replace "'") -split ",") -replace " "
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
Clear-Host
Start-O365TroubleshootersMenu
