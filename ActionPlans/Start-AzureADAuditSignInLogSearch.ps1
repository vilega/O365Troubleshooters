Function Search-AzureAdSignInAudit {
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string[]][Parameter(Mandatory=$false)] $CmdletsToSearch,
        [string][Parameter(Mandatory=$false)] $Upn)
    
        $DaysToSearch=10
        $Upn ="admin@vilega.onmicrosoft.com"
        Get-AzureADAuditSignInLogs -Filter "userPrincipalName eq $Upn"

        if (!([string]::IsNullOrEmpty($userIds)))
        {
            $UnifiedAuditLogs =Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch -UserIds $userIds -SessionCommand ReturnLargeSet 
        }
        else
        {
            $UnifiedAuditLogs =Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch  -SessionCommand ReturnLargeSet 
        }
      
        return AzureAdSignInAudit
}

$Workloads = "AzureADPreview"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\ExchangeOnlineAudit_$ts"
mkdir $ExportPath -Force

do
{
    Write-Host "Please imput the number of days you want to search (maximum 90): " -ForegroundColor Cyan -NoNewline
    [int]$DaysToSearch= Read-Host
} while ($DaysToSearch -gt 90)

Write-Host "Please imput cmdlets to search separated by comma (or just hit [Enter] to look for all cmdles): " -ForegroundColor Cyan -NoNewline
$CmdletsToSearch = Read-Host
Write-Host "Please imput the UPN for the user you want to search actions (or just hit [Enter] to look for all users): " -ForegroundColor Cyan -NoNewline
$Caller = Read-Host

$AuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToSearch -CmdletsToSearch  $CmdletsToSearch -Caller $Caller
$AuditLogs | Export-Csv "$ExportPath\ExchangeOnlineAudit_$ts.csv"
Write-Host "Exchange Online audit logs have been exported to: $ExportPath\ExchangeOnlineAudit_$ts.csv"
Read-Key

# Return to the main menu
Clear-Host
Start-O365TroubleshootersMenu


$MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
. "$MFAExchangeModule" |Out-Null
Connect-EXOPSSession -userprincipalname "admin@vilega.onmicrosoft.com"


Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName |`
where { $_ -notmatch "_none_" } | select -First 1)

$EXOSession = New-ExoPSSession -UserPrincipalName $UPN
Import-PSSession $EXOSession -AllowClobber |out-null