Clear-Host

##Connect to EXO Module via Connect-O365PS from O365Troubleshooters.psm1
$Workloads = "exo"
Connect-O365PS $Workloads

#Check if log path already exists before creating folder
if(!(Test-Path "$global:WSPath\ExportQuarantineMessage"))
{
    Clear-Host

    Write-Host "Created Log folder" -ForegroundColor Green
    mkdir -Path $global:WSPath\ExportQuarantineMessage
    Read-Host "Press ENTER to Continue"
}

Clear-Host

##Collects an Array of Quarantine Messages
##Write Archive file with all Quarantine Message EMLs to LogPath
$i = 1

$QuarantineMessages = @(Get-QuarantineMessage | Out-GridView -PassThru)

if($QuarantineMessages.Count -ne 0)
{
    foreach($QuarantineMessage in $QuarantineMessages)
    {   
        Write-Host "Exporting Quarantine Message #$i" -ForegroundColor Green
        
        try{
        $ExportedQuarantineMessage = Export-QuarantineMessage -Identity $QuarantineMessage.Identity
    
        $QuarantineMessageBytes = [Convert]::FromBase64String($ExportedQuarantineMessage.Eml)
    
        $QuarantineMessagePath = "$global:WSPath\ExportQuarantineMessage\"+$QuarantineMessage.Identity.Split('\')[1]+".eml"

        [System.IO.File]::WriteAllBytes($QuarantineMessagePath,$QuarantineMessageBytes)

        Compress-Archive -Path $QuarantineMessagePath -Update -CompressionLevel Optimal `
            -DestinationPath "$global:WSPath\ExportQuarantineMessage\QuarantineMessages.zip"

        Remove-Item $QuarantineMessagePath -Force
        }
        catch{
            Write-Log -function Export-QuarantineMessage -step ExportQuarantineMessage `
            -Description "Could Export/Write/Archive/Purge EML with`r`n"+$Error.Exception.Message
        }
        $i++
        
        Start-Sleep -s 0.5
        
        Clear-Host        
    }
    

    
    Write-Host "Created Archive with Exported Quarantine Messages 
$global:WSPath\ExportQuarantineMessage\QuarantineMessages.zip" -ForegroundColor Green
    
    Read-Host "Press ENTER to return to O365Troubleshooters Main Menu"

    Clear-Host
    
    Disconnect-All

    Start-O365TroubleshootersMenu
}

else
{
    Read-Host "No Messages were selected
Press ENTER to return to O365Troubleshooters Main Menu"

    Clear-Host

    Disconnect-All

    Start-O365TroubleshootersMenu
}