<#
        .SYNOPSIS
        Transform the value recieved from IMCEAEX NDR into an X500 that can be added as additional alias 

        .DESCRIPTION
        Provide Microsoft Defender for Office 365 Safe Links and export in a HTML format the original URL
        Can be executed on multiple encoded URL and in the end all decoded URLs can be seen the the HTML output

        .EXAMPLE
        Provide the re-written URL:
        https://nam06.safelinks.protection.outlook.com/?url=http://www.contoso.com/&data=04|01|user1@contoso.com|83ffsdfa384443fadq342743b|72f988fasdfa4d011db47|1|0|6376688415|Unknown|TWFpbGZMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwfadsfaCI6Mn0=|1000&sdata=qOwctqh5fadfaai/tglS4avTxToy67X4M8fadsfasaA=&reserved=0
        
        .LINK
        Online documentation: https://answers.microsoft.com/

    #>

Clear-Host
    
$CurrentProperty = "Collecting IMCEAEX"
$CurrentDescription = "Success"
write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\X500_$ts"
mkdir $ExportPath -Force |out-null

#TODO: ask for another IMCEAEX
$CurrentProperty = "Collecting IMCEAEX"
$CurrentDescription = ""
write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
Write-Host "`nPlease input the IMCEAEX (old LegacyExchangeDN from NDR) to transform it to X500 address: " -ForegroundColor Cyan
try 
{
    $Imceaex = Read-Host -ErrorAction Stop
}
catch 
{
    $CurrentProperty = "Collecting IMCEAEX"
    $CurrentDescription = "Error on input IMCEAEX"
    write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host "Error on input IMCEAEX, the script will return to main menu"
    Read-Key   
    Start-O365TroubleshootersMenu
}

try 
{
    # $X500 = ("X500:" + $Imceaex -replace("_","/") -replace("\+20"," ") -replace("\+28","(") -replace("\+29",")") -replace("IMCEAEX\-","") -split "@")[0] 
    $X500 = ("X500:" + $Imceaex -replace("_","/")  -replace("IMCEAEX\-","") -split "@")[0]
    $matches = ([regex]'([+][0-9a-fA-F][0-9a-fA-F])').Matches($X500)
    $HexValues = $matches | Select-Object value -Unique
    foreach($HexValue in $HexValues)
    {
        $replace = [Convert]::ToChar([Convert]::ToInt64(($HexValue.Value -replace("\+","")),16))
        $X500 = $X500 -replace("\$($HexValue.Value)",$replace)
    }
}
catch 
{
    $CurrentProperty = "Generating X500"
    $CurrentDescription = "Error on transforming $($Imceaex ) to X500"
    write-log -Function "X500FromImceaexNdr" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host "Error on transforming IMCEAEX (old LegacyExchangeDn) to X500, the script will return to main menu"
    Read-Key   
    Start-O365TroubleshootersMenu
}


#TODO: export to HTML
$ts= get-date -Format yyyyMMdd_HHmmss
$x500 |Out-File $ExportPath\x500_$ts.txt

Write-Host "The old LegacyExchangeDn was transformed to the following X500 address:"
Write-Host "$X500"
Write-Host "The script will return to main menu."
Read-Key
Start-O365TroubleshootersMenu

