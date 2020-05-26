$Workloads = "exo"
Connect-O365PS $Workloads

$i = 1

$QuarantineMessages = @(Get-QuarantineMessage | Out-GridView -PassThru)

if($QuarantineMessages.Count -ne 0)
{
    foreach($QuarantineMessage in $QuarantineMessages)
    {
        $ExportedQuarantineMessage = Export-QuarantineMessage -Identity $QuarantineMessage.Identity
    
        $QuarantineMessageBytes = [Convert]::FromBase64String($ExportedQuarantineMessage.Eml)
    
        $QuarantineMessagePath = "$global:WSPath\ExportQuarantineMessage\QuarantineMessage$i.eml"

        [System.IO.File]::WriteAllBytes($QuarantineMessagePath,$QuarantineMessageBytes)

        Write-Host "Exported $QuarantineMessagePath" -ForegroundColor Green
    
        Start-Sleep -s 0.5

        $i++
    }
    
    Compress-Archive -Path $global:WSPath\ExportQuarantineMessage\QuarantineMessage* `
            -DestinationPath "$global:WSPath\ExportQuarantineMessage\QuarantineMessages.zip" -CompressionLevel Optimal
    
    Write-Host "Created Archive with Exported Quarantine Messages $global:WSPath\ExportQuarantineMessage\QuarantineMessages.zip" `
                    -ForegroundColor Green
    
    Read-Host "Press any key to return to O365Troubleshooters Main Menu"

    Clear-Host

    Start-O365TroubleshootersMenu
}

else
{
    Read-Host "No Messages were selected
Press any key to return to O365Troubleshooters Main Menu"
    Clear-Host
    Start-O365TroubleshootersMenu
}