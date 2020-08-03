#ToDo move Main to Function to . source in other APs
Function Search-AzureAdSignInAudit {
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string][Parameter(Mandatory=$false)] $Upn)
    
        $startD = ((Get-Date).addDays(-$DaysToSearch)) 
        $startDate = "$($startD.Year)-$($startD.Month)-$($startD.Day)"
        $endD = Get-Date 
        $endDate = "$($endD.Year)-$($endD.Month)-$($endD.Day)"

        if ([string]::IsNullOrEmpty($upn))
        {
            $filterAll = "createdDateTime ge $startDate and createdDateTime le $endDate"
            $filterFail = "createdDateTime ge $startDate and createdDateTime le $endDate and status/errorCode ne 0"
            $global:AzureAdSignInAll = Get-AzureADAuditSignInLogs -Filter $filterAll
            $global:AzureAdSignInFail = Get-AzureADAuditSignInLogs -Filter $filterFail
        }
        else
        {
            $filterAll = "userPrincipalName eq `'$Upn`' and createdDateTime ge $startDate and createdDateTime le $endDate"
            $filterFail = "userPrincipalName eq `'$Upn`' and createdDateTime ge $startDate and createdDateTime le $endDate and status/errorCode ne 0"
            $global:AzureAdSignInAll = Get-AzureADAuditSignInLogs -Filter $filterAll
            $global:AzureAdSignInFail = Get-AzureADAuditSignInLogs -Filter $filterFail
        }
      
}

Function Start-AzureADAuditSignInLogSearch
{
    Clear-Host
    $Workloads = "AzureADPreview"
    Connect-O365PS $Workloads

    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

    Write-Host "Retrieving sign in logs is based on a preview feature!`n" -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    $ts= get-date -Format yyyyMMdd_HHmmss
    $ExportPath = "$global:WSPath\AzureADSignInAudit_$ts"
    mkdir $ExportPath -Force |out-null

    Write-Warning "Please be aware AzureAD Sign In logs availability is limited`r`nFor Azure AD Free you can retrieve 7 days`r`nFor Azure AD Premium P1/P2 you can retrieve 30days"

    <#do
    {
        Write-Host "Please input the number of days you want to search (maximum 90): " -ForegroundColor Cyan -NoNewline
    } while ($DaysToSearch -gt 90)#>

    [int]$DaysToSearch= Read-IntFromConsole -IntType "Number of days to investigate Azure AD Sign In Logs"

    if($DaysToSearch -gt 30)
    {
        Write-Warning "We will only be able to provide a maximum of 30 days for this log"
        [int]$DaysToSearch = 30
    }

    Write-Host "Please input the UPN for the user you want to search sign in logs (or just hit [Enter] to look for all users): " -ForegroundColor Cyan -NoNewline
    $Upn = Read-Host

    Search-AzureAdSignInAudit -DaysToSearch $DaysToSearch -Upn $Upn
    $global:AzureAdSignInAll | Export-Csv "$ExportPath\AllSignInAuditLogs_$ts.csv" -NoTypeInformation
    $global:AzureAdSignInFail | Export-Csv "$ExportPath\FailSignInAuditLogs_$ts.csv" -NoTypeInformation
    Write-Host "Azure AD sign in logs (all and fail) have been exported to: $ExportPath" -ForegroundColor Green
    Read-Key

    # Return to the main menu
    Clear-Host
    Start-O365TroubleshootersMenu
}