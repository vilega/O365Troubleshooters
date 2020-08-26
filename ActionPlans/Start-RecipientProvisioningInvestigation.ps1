
Function Search-RecipientObject
{
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string[]][Parameter(Mandatory=$false)] $OperationsToSearch,
        [string][Parameter(Mandatory=$false)] $userIds)
      
    $DaysToSearch=10
    if (!([string]::IsNullOrEmpty($userIds)))
    {
        $UnifiedAuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch - -UserIds $userIds -SessionCommand ReturnLargeSet 
    }
    else
    {
        $UnifiedAuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch  -SessionCommand ReturnLargeSet 
    }
  
    return $UnifiedAuditLogs

}


Clear-Host
$Workloads = "exo", "msol"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RecipientProvisioningInvestigation_$ts"
mkdir $ExportPath -Force |out-null
