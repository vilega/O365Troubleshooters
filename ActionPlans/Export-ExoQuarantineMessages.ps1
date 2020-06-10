Clear-Host

##Connect to EXO Module via Connect-O365PS from O365Troubleshooters.psm1
$Workloads = "exo"
Connect-O365PS $Workloads

$ts= get-date -Format yyyyMMdd_HHmmss
#Check if log path already exists before creating folder
if(!(Test-Path "$global:WSPath\ExportQuarantineMessage_$ts"))
{
    Clear-Host

    $QuarantineMessageExportPath = "$global:WSPath\ExportQuarantineMessage_$ts"
    
    mkdir -Path $QuarantineMessageExportPath | Out-Null
    
    Write-Host "Created Log folder`r`n$QuarantineMessageExportPath" -ForegroundColor Green

    Read-Key
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
    
        $QuarantineMessagePath = $QuarantineMessageExportPath+"\"+$QuarantineMessage.Identity.Split('\')[1]+".eml"

        [System.IO.File]::WriteAllBytes($QuarantineMessagePath,$QuarantineMessageBytes)

        Compress-Archive -Path $QuarantineMessagePath -Update -CompressionLevel Optimal `
            -DestinationPath "$QuarantineMessageExportPath\QuarantineMessages.zip"

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
$QuarantineMessageExportPath\QuarantineMessages.zip
You will be returned to O365Troubleshooters Main Menu" -ForegroundColor Green
    
    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}

else
{
    Write-Host "No Messages were selected, you will be returned to O365Troubleshooters Main Menu" -ForegroundColor Red

    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}