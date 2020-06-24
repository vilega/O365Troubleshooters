Function Start-CompromisedMain
{

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



}



#region Connectors Created
# Inbound Connector Check
$InboundConnectorsCollection = @()
$InboundConnectors = Get-InboundConnector | Where-Object ConnectorType -EQ "OnPremises"
$now = (Get-date).ToUniversalTime()
#([datetime]::UtcNow)
$DaysToInvestigate = 14

foreach($InboundConnector in $InboundConnectors) {

    $ts = New-TimeSpan -Start $InboundConnector.WhenChangedUTC -End $now
    #$InboundConnectorCollection += $InboundConnector |select  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
    $InboundConnectorsCollection += $InboundConnector |Select-Object  *, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
 
}
Write-Host "The following Inbound On Premises connectors have been changed/created in the last 14 days" -ForegroundColor Red
foreach ($InboundConnectorCollection in $InboundConnectorsCollection) {
    if ($InboundConnectorCollection.DaysSinceLastChange -le $DaysToInvestigate)
        {
            $InboundConnectorCollection |Select-Object  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, DaysSinceLastChange
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
    $OutboundConnectorsCollection += $OutboundConnector |Select-Object  *, @{Name='DaysSinceLastChange';Expression={$ts.Days}}
 
}
Write-Host "The following Outbound On Premises connectors have been changed/created in the last 14 days" -ForegroundColor Red
foreach ($OutboundConnectorCollection in $OutboundConnectorsCollection) {
    if ($OutboundConnectorCollection.DaysSinceLastChange -le $DaysToInvestigate)
        {
            $OutboundConnectorCollection |Select-Object  Name,SenderDomains,TlsSenderCertificateName, SenderIPAddresses, DaysSinceLastChange
        }
}

#ToDo
#Here we will use Search-EXOAdminAudit from EXO Audit Search AP
#We will need to modify EXO Audit Search AP so we can dot source and re-use the function we want.
$AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch 900 -CmdletsToSearch "New-InboundConnector","Set-InboundConnector","New-OutboundConnector","Set-OutboundConnector","Remove-InboundConnector","Remove-OutboundConnector"

#endregion Connectors Created


<#
Transport Rules - Done
Forwarding
Redirect - Done
Journaling - Done
CBR - Done
BCC - Done
Audit 14
#>

#region TransportRules

<#$SuspiciousTransportRule = New-Object -TypeName psobject
$SuspiciousTransportRule | Add-Member -MemberType NoteProperty -Name RuleType
$SuspiciousTransportRule | Add-Member -MemberType NoteProperty -Name Description#>
$SuspiciousTransportRules = @()
$TransportRules = Get-TransportRule -ResultSize unlimited
foreach($TransportRule in $TransportRules)
{
    switch -wildcard($TransportRule.Description)
    {   
        "*redirect the message to*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        "*Route the message using the connector*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        "*Blind carbon copy(Bcc) the message*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        "*Forward the message*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        "*Add the sender's manager as recipient type*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        "*Send the incident report to*" 
            {$SuspiciousTransportRules += $TransportRule|Select-Object Name,Description,State,Guid,WhenChanged;break}
        default{}
    }
}
<#
(Get-TransportRule -Filter "Description -like '*redirect the message to*'").Description
(Get-TransportRule -Filter "Description -like '*Route the message using the connector*'").Description
(Get-TransportRule -Filter "Description -like '*Blind carbon copy(Bcc) the message*'").Description
(Get-TransportRule -Filter "Description -like '*Forward the message*'").Description
(Get-TransportRule -Filter "Description -like '*Add the sender's manager as recipient type*'").Description
(Get-TransportRule -Filter "Description -like '*Send the incident report to*'").Description 
#>


$AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate -CmdletsToSearch "New-TransportRule","Set-TransportRule","Remove-TransportRule"

#endregion TransportRules

#region JournalRule
$JournalRule = @()
$JournalRule = Get-JournalRule
if($JournalRule.count -eq 0)
{
    Write-Host "No Journal Rule"
}
else 
{
    Write-Host "We have detected the following Journal Rules:"
    $JournalRule|Format-Table Identity, Enabled, Scope, JournalEmailAddress, Recipient, WhenChanged
}
#endregion JournalRule


#region Check GA
$Administrators = Get-MsolRole | ForEach-Object{if (($_.name -eq "Company Administrator") -or ($_.name -eq "Exchange Service Administrator")) {$_}} |ForEach-Object{Get-MsolRoleMember -MaxResults 10000 -RoleObjectId $_.ObjectID}
$AdministratorsList = @()
foreach($Administrator in $Administrators)
{

    $MsolUser = get-msoluser -UserPrincipalName $Administrator.EmailAddress |Select-Object LastPasswordChangeTimestamp, StrongPasswordRequired
    $mailbox = get-mailbox $Administrator.EmailAddress -ErrorAction SilentlyContinue |Select-Object ForwardingAddress,ForwardingSmtpAddress, DeliverToMailboxAndForward
    
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
$AdministratorsList |Format-Table


Get-InboxRule -Mailbox $Upn|Format-List

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