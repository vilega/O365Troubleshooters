<#
        .SYNOPSIS
        Decode Microsoft Defender for Office 365 Safe Links to show original URL 

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

# Variable to know if any URL needs to be decoded
[bool]$decode = $true 

# Create timestamp
$ts = get-date -Format yyyyMMdd_HHmmss

# Create export folder
try {
    $ExportPath = "$global:WSPath\DecodeSafeLinksUrl_$ts"
    mkdir $ExportPath -Force | out-null
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Create ExportPath" -Description "Success"
}
catch {
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Create ExportPath" -Description "Couldn't create folder $global:WSPath\DecodeSafeLinksUrl_$ts. Error: $($_.Exception.Message)"
    Write-Host "Couldn't create folder $global:WSPath\DecodeSafeLinksUrl_$ts"
    Read-Key
    Start-O365TroubleshootersMenu
}



#Import assembly System.Web which contains HttpUtility.UrlDecode method
try {
    Add-Type -AssemblyName System.Web
}
catch {
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Referrence assembly System.Web" -Description "Error: $($_.Exception.Message)"
    Write-Host "Couldn't load assembly System.Web. The script will return to the main menu"
    Read-Key
    Start-O365TroubleshootersMenu
}


# creating a list to store original URLs
$ListOfOriginalAndDecodedUrls = New-Object -TypeName "System.Collections.ArrayList"

While ($decode) {
    # Read from console the encoded URL
    $encodedURL = Read-Host("Please provide the Microsoft Defender for Office 365 Safe Links URL that you want to decode to original URL")
  
    
    

    try {   
        # Decode URL using UrlDecode from System.Web.HttpUtility
        $decodedURL = [System.Web.HttpUtility]::UrlDecode($encodedURL)
    
    
        #$decodedURL = (($decodedURL -Split "url=")[1] -split "&data=;")[0]
    
        # check if decoded URL is of SafeLinks format
        # throw System.ArgumentException if the format is not supported
        if ($decodedURL -match ".safelinks.protection.outlook.com\/.*\?url=.+&data=") {
            $decodedURL = (($decodedURL -Split "/?url=")[1] -Split "&data=")[0]
        }
        elseif ($decodedURL -match ".safelinks.protection.outlook.com\/.*\?url=.+&amp;data=") {
            $decodedURL = (($decodedURL -Split "/?url=")[1] -Split "&amp;data=")[0]
        }
        else { throw  New-Object System.ArgumentException "$encodedURL is not in the correct Safe Links format", "encodedURL" }
    }
    
    catch [System.ArgumentException] {
        Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Couldn't decode and parse URL: $encodedURL"
        Write-Host "Couldn't decode and parse URL: $encodedURL"
        $decodedURL = $null
        Read-Key
    }
    catch {
        Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Unhandled error! Input URL: $encodedURL, Exception message: $($PSItem.Exception.Message)"
        Write-Host "Unhandled error! Input URL: $encodedURL, Exception message: $PSItem.Exception.Message"
        $decodedURL = $null
        Read-Key
    }

    # Log at the console and logging file the decoded URL
    Write-Host "The decoded URL is:" -ForegroundColor Green
    Write-Host $decodedURL
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Decoded and Parse URL is: $decodedURL"
    Read-Key
    $urlHashTabel = @{
        encodedURL = $encodedURL
        decodedURL = $decodedURL 
    }

    #Cast HashTabel into PSCustomObject and add it to the List collection
    $null = $ListOfOriginalAndDecodedUrls.Add([PSCustomObject]$urlHashTabel)
    
    # Ask if any new URL needs to be decoded
    Clear-Host
    Write-Host "Do you need a new URL to be decoded?"
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
    for ($i = 0; $i -lt $ListOfOriginalAndDecodedUrls.Count; $i++) {
        if ($null -eq $ListOfOriginalAndDecodedUrls[$i].decodedURL) {
            [string]$SectionTitle = "Decode Safe Links URL - $($i+1)"
            [string]$Description = "$encodedURL is not in the correct Safe Links format"
            [PSCustomObject]$ListOfOriginalAndDecodedUrlsHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString " "
            $null = $TheObjectToConvertToHTML.Add($ListOfOriginalAndDecodedUrlsHtml)
        }
        else {
            [string]$SectionTitle = "Decode Safe Links URL - $($i+1)"
            [string]$Description = "The encoded Microsoft Defender for Office 365 Safe Links URL is decoded to show the original URL"
            [PSCustomObject]$ListOfOriginalAndDecodedUrlsHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList  $ListOfOriginalAndDecodedUrls[$i] -TableType "List"
            $null = $TheObjectToConvertToHTML.Add($ListOfOriginalAndDecodedUrlsHtml)
        }

    }

    #Build HTML report out of the previous HTML sections
    [string]$FilePath = $ExportPath + "\DecodeSafeLinksUrl.html"
    Export-ReportToHTML -FilePath $FilePath -PageTitle "Microsoft Defender for Office 365 Safe Links Decoder" -ReportTitle "Microsoft Defender for Office 365 Safe Links Decoder" -TheObjectToConvertToHTML $TheObjectToConvertToHTML

    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Generate HTML Report" -Description "Success"
}
catch {
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Generate HTML Report" -Description "Error: $($PSItem.Exception.Message)"
    
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
    $ListOfOriginalAndDecodedUrls | Export-Csv -Path "$ExportPath\DecodeSafeLinksUrl.csv" -NoTypeInformation
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Generate CSV Report" -Description "Success"
}
catch {
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Generate CSV Report" -Description "Error: $($PSItem.Exception.Message)"
}

# return to main menu
Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Load Start-O365Troubleshooters Menu" -Description "Success"
Write-Host "The script will return to main menu."
Read-Key
Start-O365TroubleshootersMenu