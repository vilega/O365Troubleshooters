Clear-Host
$encodedURL  = Read-Host("Please provide the ATP SafeLinks URL that you want to decode to original URL")
Add-Type -AssemblyName System.Web

try
{
    $decodedURL = [System.Web.HttpUtility]::UrlDecode($encodedURL)
    #$decodedURL = (($decodedURL -Split "url=")[1] -split "&data=;")[0]
    if($decodedURL -match "url=(\S+)&data=\S+"){$decodedURL = $Matches[1]}
    elseif($decodedURL -match "url=(\S+)&amp;data"){$decodedURL = $Matches[1]}
}
catch
{
    Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Couldn't decode and parse URL: $encodedURL"
    Write-Host "Couldn't decode and parse URL: $encodedURL"
    Read-Host "Press any key and then to reload main menu [Enter]"
    Start-O365TroubleshootersMenu
}

Write-Host "The decoded URL is:" -ForegroundColor Green
Write-Host $decodedURL
Write-Log -function "Start-AP_DecodeSafeLinksURL" -step  "Decoding URL" -Description "Decoded and Parse URL is: $decodedURL"
Read-Key
Start-O365TroubleshootersMenu