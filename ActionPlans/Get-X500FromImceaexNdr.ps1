<#
        .SYNOPSIS
        Transform the value recieved from IMCEAEX NDR into an X500 that can be added as additional alias 

        .DESCRIPTION
        When you send email messages to an internal user in Microsoft Office 365, you receive an IMCEAEX non-delivery report (NDR) because of a bad LegacyExchangeDN reference. The IMCEAEX NDR indicates that the user no longer exists in the environment.
        The auto-complete cache in Microsoft Outlook and in Microsoft Outlook Web App (OWA) uses the value of the LegacyExchangeDN attribute to route email messages internally.
        To resolve this issue you need to create an X500 proxy address for the old LegacyExchangeDN.
        Out tool will help you to generate the exact X500 you need to add as an aditional alias on the affected recipient
        https://docs.microsoft.com/exchange/troubleshoot/email-delivery/imceaex-ndr

        .EXAMPLE
        You can see more details on: https://docs.microsoft.com/exchange/troubleshoot/email-delivery/imceaex-ndr        
        
        .LINK
        Online documentation: https://aka.ms/O365Troubleshooters/GenerateX500FromImceaexNDR

    #>

Clear-Host
    
$CurrentProperty = "Collecting IMCEAEX"
$CurrentDescription = "Start"
write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 

# Create timestamp
$ts = get-date -Format yyyyMMdd_HHmmss

# Create export folder
try {
    $ExportPath = "$global:WSPath\X500_$ts"
    mkdir $ExportPath -Force | out-null
}
catch {
    Write-Log -function "Get-X500FromImceaexNdr" -step  "create export folder" -Description "Error: $($_.Exception.Message)"
    Write-Host "Couldn't create the export folder. The script will return to the main menu"
    Read-Key
    Start-O365TroubleshootersMenu
}


# creating a list to store original URLs
$ListOfOriginalImceaexAndX500 = New-Object -TypeName "System.Collections.ArrayList"

# Loop until will be false (when administrator won't continue)
$decode = $true

While ($decode) {
   
    #ask for IMCEAEX
    $CurrentProperty = "Collecting IMCEAEX"
    $CurrentDescription = ""
    write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host "`nPlease input the IMCEAEX (old LegacyExchangeDN from NDR) to transform it to X500 address: " -ForegroundColor Cyan
    try {
        $Imceaex = Read-Host -ErrorAction Stop
    }
    catch {
        $CurrentProperty = "Collecting IMCEAEX"
        $CurrentDescription = "Error on input IMCEAEX"
        write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
        Write-Host "Error on input IMCEAEX, the script will return to main menu"
        Read-Key   
        Start-O365TroubleshootersMenu
    }

    try {
        # $X500 = ("X500:" + $Imceaex -replace("_","/") -replace("\+20"," ") -replace("\+28","(") -replace("\+29",")") -replace("IMCEAEX\-","") -split "@")[0] 
        $X500 = ("X500:" + $Imceaex -replace ("_", "/") -replace ("IMCEAEX\-", "") -split "@")[0]
        $matches = ([regex]'([+][0-9a-fA-F][0-9a-fA-F])').Matches($X500)
        $HexValues = $matches | Select-Object value -Unique
        foreach ($HexValue in $HexValues) {
            $replace = [Convert]::ToChar([Convert]::ToInt64(($HexValue.Value -replace ("\+", "")), 16))
            $X500 = $X500 -replace ("\$($HexValue.Value)", $replace)
        }
    }
    catch {
        $CurrentProperty = "Generating X500"
        $CurrentDescription = "Error on transforming $($Imceaex ) to X500"
        write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
        Write-Host "Error on transforming IMCEAEX (old LegacyExchangeDn) to X500, the script will return to main menu"
        Read-Key   
        Start-O365TroubleshootersMenu
    }

    # Log at the console and logging file the decoded URL
    Write-Host "The old LegacyExchangeDn was transformed to the following X500 address:" -ForegroundColor Green
    Write-Host $X500
    Write-Log -function "Get-X500FromImceaexNdr" -step  "Transform to X500" -Description "Decoded X500 is: $X500"
    Read-Key
    $X500HashTabel = @{
        IMCEAEX = $Imceaex
        X500    = $X500 
    }
    
    #Cast HashTabel into PSCustomObject and add it to the List collection
    $null = $ListOfOriginalImceaexAndX500.Add([PSCustomObject]$X500HashTabel)

    # Ask if any new URL needs to be decoded
    Clear-Host
    Write-Host "Do you need to convert another IMCEAEX (old LegacyExchangeDN from NDR)?"
    $answer = Get-Choice "Yes", "No"
    if ($answer -eq 'y') {
        $decode = $true 
    }
    elseif ($answer -eq 'n') {
        $decode = $false 
    }

}





#region CreateHtmlReport

try {
    

    #Create the collection of sections of HTML
    $TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"
    for ($i = 0; $i -lt $ListOfOriginalImceaexAndX500.Count; $i++) {

            [string]$SectionTitle = "Decoded X500 - $($i+1)"
            [string]$Description = "The IMCEAEX NDR is decoded to create the X500"
            [PSCustomObject]$ListOfImceaexAndX500Html = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList  $ListOfOriginalImceaexAndX500[$i] -TableType "List"
            $null = $TheObjectToConvertToHTML.Add($ListOfImceaexAndX500Html)
    }

    #Build HTML report out of the previous HTML sections
    [string]$FilePath = $ExportPath + "\X500.html"
    Export-ReportToHTML -FilePath $FilePath -PageTitle "The IMCEAEX NDR is decoded to create the X500" -ReportTitle "The IMCEAEX NDR is decoded to create the X500" -TheObjectToConvertToHTML $TheObjectToConvertToHTML

    Write-Log -function "Get-X500FromImceaexNdr" -step  "Generate HTML Report" -Description "Success"
}
catch {
    Write-Log -function "Get-X500FromImceaexNdr" -step  "Generate HTML Report" -Description "Error: $($PSItem.Exception.Message)"
    
}
#Ask end-user for opening the HTMl report
$OpenHTMLfile = Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile.ToLower() -like "*y*") {
    Write-Host "Opening report...." -ForegroundColor Cyan
    Start-Process $FilePath
}
#endregion ResultReport
   
# Print location where the data was exported
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Read-Key


# Create CSV 
try {
    $ListOfOriginalAndDecodedUrls | Export-Csv -Path "$ExportPath\X500.csv" -NoTypeInformation
    Write-Log -function "Get-X500FromImceaexNdr" -step  "Generate CSV Report" -Description "Success"
}
catch {
    Write-Log -function "Get-X500FromImceaexNdr" -step  "Generate CSV Report" -Description "Error: $($PSItem.Exception.Message)"
}

# return to main menu
Write-Log -function "Get-X500FromImceaexNdr" -step  "Load Start-O365Troubleshooters Menu" -Description "Success"
Write-Host "The script will return to main menu."
Read-Key
Start-O365TroubleshootersMenu