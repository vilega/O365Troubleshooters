Function Search-UnifiedLog

{
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string[]][Parameter(Mandatory=$false)] $OperationsToSearch,
        [string][Parameter(Mandatory=$false)] $Caller)
      
    $DaysToSearch=10
    if (!([string]::IsNullOrEmpty($userIds)))
    {
        $UnifiedAuditLogs =Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch -UserIds $userIds -SessionCommand ReturnLargeSet 
    }
    else
    {
        $UnifiedAuditLogs =Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch  -SessionCommand ReturnLargeSet 
    }
  
    return UnifiedAuditLogs

}

$Workloads = "exo"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\UnifiedAudit_$ts"
mkdir $ExportPath -Force
do
{
    Write-Host "Please imput the number of days you want to search (maximum 90): " -ForegroundColor Cyan -NoNewline
    $DaysToSearch= Read-Host
}while ($DaysToSearch -le 90)


if ((Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled)
{#
    if (!((Get-Date).addDays(-$DaysToSearch) -ge (Get-AdminAuditLogConfig).UnifiedAuditLogFirstOptInDate))
    {
        Write-Host "Unified Audit Log is enabled but don't include all required days to search." -ForegroundColor Yellow
        #TODO: write-log
        Write-Host "Unified Audit Log has been enabled on $((Get-AdminAuditLogConfig).UnifiedAuditLogFirstOptInDate) and will contain only logs after it was enabled" -ForegroundColor Yellow
       
    }
}
else
{
    Write-Host "Unified Audit Log is disabled." -ForegroundColor Red
    Write-Host "Script returns to Main Menu"
    Read-Key
    Clear-Host
    Start-O365TroubleshootersMenu
}

#Write-Host "Please imput Operations to search separated by comma (or just hit [Enter] to look for all cmdles): " -ForegroundColor Cyan -NoNewline
#$Operations = Read-Host
Write-Host "Please imput the UPN for the user you want to search actions (or just hit [Enter] to look for all users): " -ForegroundColor Cyan -NoNewline
$userIds = Read-Host

Search-UnifiedLog

$UnifiedAuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToSearch -OperationsToSearch  $Operations -UserIds $UserIds
$UnifiedAuditLogs | Export-Csv "$ExportPath\ExchangeOnlineAudit_$ts.csv"
Write-Host "Exchange Online audit logs have been exported to: $ExportPath\ExchangeOnlineAudit_$ts.csv"
Write-Host "To parse and use the generated audit logs, go to the article: https://docs.microsoft.com/en-us/microsoft-365/compliance/export-view-audit-log-records ."
Read-Key


# Return to the main menu
Clear-Host
Start-O365TroubleshootersMenu


###
$LiveCred = Get-Credential
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/  -Credential $LiveCred -Authentication Basic -AllowRedirection
$ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session -AllowClobber
Import-PSSession $ccSession -Prefix cc -AllowClobber

