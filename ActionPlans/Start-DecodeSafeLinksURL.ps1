#.SYN...


#
Clear-Host

# 
$encodedURL  = Read-Host("Please provide the ATP SafeLinks URL that you want to decode to original URL")
Add-Type -AssemblyName System.Web


try
{   # Decode URL using UrlDecode from System.Web.HttpUtility
    $decodedURL = [System.Web.HttpUtility]::UrlDecode($encodedURL)
    
    
    #$decodedURL = (($decodedURL -Split "url=")[1] -split "&data=;")[0]
    
    # 
    if($decodedURL -match ".safelinks.protection.outlook.com\/.*\?url=.+&data=")
    {
        $decodedURL = (($decodedURL -Split "/?url=")[1] -Split "&data=")[0]
    }
    elseif($decodedURL -match ".safelinks.protection.outlook.com\/.*\?url=.+&amp;data=")
    {
        $decodedURL = (($decodedURL -Split "/?url=")[1] -Split "&amp;data=")[0]
    }
    else{throw "InvalidSafeLinksURL"}
}
#
catch [InvalidSafeLinksURL]
{
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Couldn't decode and parse URL: $encodedURL"
    Write-Host "Couldn't decode and parse URL: $encodedURL"
    Read-Host "Press any key and then to reload main menu [Enter]"
    Start-O365TroubleshootersMenu
}
catch
{

}

# ask for new decode

# HTML & TXT
Write-Host "The decoded URL is:" -ForegroundColor Green
Write-Host $decodedURL
Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Decoded and Parse URL is: $decodedURL"
Read-Key
Start-O365TroubleshootersMenu