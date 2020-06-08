Function Search-AzureAdSignInAudit {
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string][Parameter(Mandatory=$false)] $Upn)
    
        $startD = ((Get-Date).addDays(-$DaysToSearch)) 
        $startDate = "$($startD.Year)-$($startD.Month)-$($startD.Day)"
        $endD = Get-Date 
        $endDate = "$($endD.Year)-$($endD.Month)-$($endD.Day)"

        if (!([string]::IsNullOrEmpty($userIds)))
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

do
{
    Write-Host "Please input the number of days you want to search (maximum 90): " -ForegroundColor Cyan -NoNewline
    [int]$DaysToSearch= Read-Host
} while ($DaysToSearch -gt 90)


Write-Host "Please input the UPN for the user you want to search sign in logs (or just hit [Enter] to look for all users): " -ForegroundColor Cyan -NoNewline
$Upn = Read-Host

Search-AzureAdSignInAudit -DaysToSearch $DaysToSearch -Upn $Upn
$global:AzureAdSignInAll | Export-Csv "$ExportPath\AllSignInAuditLogs_$ts.csv" -NoTypeInformation
$global:AzureAdSignInFail | Export-Csv "$ExportPath\FailSignInAuditLogs_$ts.csv" -NoTypeInformation
Write-Host "Azure AD sign in logs (all and fail) have been exported to: $ExportPath"
Read-Key

# Return to the main menu
Clear-Host
Start-O365TroubleshootersMenu
