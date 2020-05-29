Function Start-CompromisedMain
{


}

#region Blocked Senders
###Get Blocked Senders and Create Hashtable Array with SenderAddress & Reasons
$blockedSenders = Get-BlockedSenderAddress
$Reasons = @()

 

foreach($blockedSender in $blockedSenders)
{
    $Reason = $blockedSender.Reason.Replace(";","`n")
    $Reason = ConvertFrom-StringData $Reason.Replace(":","=")
    $Reason["SenderAddress"] = $blockedSender.SenderAddress
    $Reasons += $Reason
}

$Reasons
#endregion Blocked Senders

#region Connectors Created
$InboundConnectorsCollection = @()
$InboundConnectors = Get-InboundConnector | ? ConnectorType -EQ "OnPremises"
$now = (Get-date).ToUniversalTime()
#([datetime]::UtcNow)
$DaysToInvestigate = 14

foreach($InboundConnector in $InboundConnectors) {

    $ts = New-TimeSpan -Start $InboundConnector.WhenChangedUTC -End $now
    #$InboundConnectorCollection += $InboundConnector |select  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
    $InboundConnectorsCollection += $InboundConnector |select  *, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
 
}
Write-Host "The following Inbound On Premises connectors have been changed/created in the last 14 days" -ForegroundColor Red
foreach ($InboundConnectorCollection in $InboundConnectorsCollection) {
    if ($InboundConnectorCollection.DaysSinceLastChange -le $DaysToInvestigate)
        {
            $InboundConnectorCollection |select  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, DaysSinceLastChange
        }
}

# Outbound Connectors check
$OutboundConnectorsCollection = @()
$OutboundConnectors = Get-OutboundConnector 
$now = (Get-date).ToUniversalTime()
#([datetime]::UtcNow)

foreach($OutboundConnector in $OutboundConnectors) {

    $ts = New-TimeSpan -Start $OutboundConnector.WhenChangedUTC -End $now
    #$OutboundConnectorCollection += $OutboundConnector |select  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
    $OutboundConnectorsCollection += $OutboundConnector |select  *, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
 
}
Write-Host "The following Outbound On Premises connectors have been changed/created in the last 14 days" -ForegroundColor Red
foreach ($OutboundConnectorCollection in $OutboundConnectorsCollection) {
    if ($OutboundConnectorCollection.DaysSinceLastChange -le $DaysToInvestigate)
        {
            $OutboundConnectorCollection |select  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, DaysSinceLastChange
        }
}

$AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch 900 -CmdletsToSearch "New-InboundConnector","Set-InboundConnector","New-OutboundConnector","Set-OutboundConnector","Remove-InboundConnector","Remove-OutboundConnector"


#endregion Connectors Created

#region TransportRules
    <#
    Transport Rules
        Forwarding
        Redirect
        Journaling
        CBR
        BCC

    Audit 14

    #>

    (Get-TransportRule -Filter "Description -like '*redirect the message to*'").Description
    (Get-TransportRule -Filter "Description -like '*Route the message using the connector*'").Description
    (Get-TransportRule -Filter "Description -like '*Blind carbon copy(Bcc) the message*'").Description

    $AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate -CmdletsToSearch "New-TransportRule","Set-TransportRule","Remove-TransportRule"


#endregion TransportRules

#region Check GA
$Administrators = Get-MsolRole | %{if (($_.name -eq "Company Administrator") -or ($_.name -eq "Exchange Service Administrator")) {$_}} |%{Get-MsolRoleMember -MaxResults 10000 -RoleObjectId $_.ObjectID}
$AdministratorsList = @()
foreach($Administrator in $Administrators)
{

    $MsolUser = get-msoluser -UserPrincipalName $Administrator.EmailAddress |select LastPasswordChangeTimestamp, StrongPasswordRequired
    $mailbox = get-mailbox $Administrator.EmailAddress -ErrorAction SilentlyContinue |select ForwardingAddress,ForwardingSmtpAddress, DeliverToMailboxAndForward
    
    $Admin = New-Object -TypeName psobject 
    $Admin | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $Administrator.EmailAddress 
    $Admin | Add-Member -MemberType NoteProperty -Name "LastPasswordChangeTimestamp" -Value $MsolUser.LastPasswordChangeTimestamp
    $Admin | Add-Member -MemberType NoteProperty -Name "MfaState" -Value $Administrator.StrongAuthenticationRequirements.State
    $Admin | Add-Member -MemberType NoteProperty -Name "StrongPasswordRequired" -Value $MsolUser.StrongPasswordRequired
    $Admin | Add-Member -MemberType NoteProperty -Name "ForwardingAddress" -Value $mailbox.ForwardingAddress
    $Admin | Add-Member -MemberType NoteProperty -Name "ForwardingSmtpAddress" -Value $mailbox.ForwardingSmtpAddress
    $Admin | Add-Member -MemberType NoteProperty -Name "DeliverToMailboxAndForward" -Value $mailbox.DeliverToMailboxAndForward

    $AdministratorsList += $Admin
}
$AdministratorsList |ft


Get-InboxRule -Mailbox $Upn|fl

#endregion Check GA

#region GA audit disable & audit bypass

if ((Get-OrganizationConfig).AuditDisabled -eq $true)
{
    Write-Host "Automatic AuditEnabled at organization level is turned off" -ForegroundColor Red

    foreach($Administrator in $AdministratorsList.UserPrincipalName)
    {
        if ((get-mailbox $Administrator -ea SilentlyContinue).AuditEnabled -eq $false)
        {
            Write-Host "The following Global Administrator $($Administrator) has mailbox audit disabled"
        }
    
        if ((Get-MailboxAuditBypassAssociation -Identity $Administrator).AuditByPassEnabled -eq $true)
        {
            Write-Host "The following administrator's ($Administrator) actions on other mailboxes are not audited!!! " -ForegroundColor Red
        }
    }
}
else
{
    Write-Host "Automatic AuditEnabled at organization level is turned on" -ForegroundColor Green
    
    foreach($Administrator in $AdministratorsList.UserPrincipalName)
    {
        if ((Get-MailboxAuditBypassAssociation -Identity $Administrator).AuditByPassEnabled -eq $true)
        {
            Write-Host "The following administrator's ($Administrator) actions on other mailboxes are not audited!!! " -ForegroundColor Red
        }
    }
}
#endregion GA audit disable & audit bypass














$Workloads = "exo","SCC", "MSOL"#, "AAD"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\Compromised_$ts"
mkdir $ExportPath -Force

. $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1


Start-CompromisedMain