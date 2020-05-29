Function Search-EXOAdminAudit {
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string[]][Parameter(Mandatory=$false)] $CmdletsToSearch,
        [string][Parameter(Mandatory=$false)] $Caller)
    
    $user = $Caller    
    if (!($user))
    {
        $user =$null
    }

    $AdminAuditLogs = Search-AdminAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Cmdlets $CmdletsToSearch -ExternalAccess $false -UserIds $user

    $ParsedAuditLogs = @()
    foreach ($AdminAuditLog in $AdminAuditLogs)
    {
        $ParsedAuditLog = New-Object -TypeName psobject 
        $ParsedAuditLog | Add-Member -MemberType NoteProperty -Name "Caller" -Value $AdminAuditLog.Caller
        $ParsedAuditLog | Add-Member -MemberType NoteProperty -Name "ClientIP" -Value $AdminAuditLog.ClientIP
        $ParsedAuditLog | Add-Member -MemberType NoteProperty -Name "Succeeded" -Value $AdminAuditLog.Succeeded
        $ParsedAuditLog | Add-Member -MemberType NoteProperty -Name "RunDate" -Value $AdminAuditLog.RunDate
        $Cmdlet = [string]$AdminAuditLog.CmdletName
        foreach ($CmdletParameters in $AdminAuditLog.CmdletParameters)
        {
            $Cmdlet += " -$($CmdletParameters.Name) `"$($CmdletParameters.Value)`""
        }
        $ParsedAuditLog | Add-Member -MemberType NoteProperty -Name "Cmdlet" -Value $Cmdlet
        $ParsedAuditLogs += $ParsedAuditLog
    }
    return $ParsedAuditLogs
}

$Workloads = "exo"
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
