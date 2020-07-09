Function Get-BlockedSenderReasons
{
    ###Get Blocked Senders and Create Hashtable Array with SenderAddress & Reasons
    ###ToDo clarify what to return to MainMenu
    $blockedSenders = Get-BlockedSenderAddress

    if($null -ne $blockedSenders)
    {   
        $blockedSenders|Export-Csv -NoTypeInformation -Path "$ExportPath\BlockedOutboundSenders.csv"
        
        $blockedSenderReasons = @()

        foreach($blockedSender in $blockedSenders)
        {
            $Reason = $blockedSender.Reason.Replace(";","`n")
            $Reason = ConvertFrom-StringData $Reason.Replace(":","=")
            $Reason["SenderAddress"] = $blockedSender.SenderAddress
            $blockedSenderReasons += $Reason
        }
        return $blockedSenderReasons
    }
    else 
    {
        Write-Host -ForegroundColor Green "No Banned Outbound Senders Found"
        return $null
    }
}


Function Get-RecentSuspiciousConnectors
{param([int][Parameter(Mandatory=$true)] $DaysToInvestigate, [datetime][Parameter(Mandatory=$true)] $CurrentDateTime)
    $InboundConnectors = Get-InboundConnector | Where-Object {($_.ConnectorType -EQ "OnPremises") -and `
            ( ($_.WhenCreatedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) -or ($_.WhenChangedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) )}

    $OutboundConnectors = Get-OutboundConnector | Where-Object {($_.WhenCreatedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) `
                                                                -or ($_.WhenChangedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate))}
    
    if($null -ne $InboundConnectors)
    {
        Write-Host "Inbound On Premises connectors have been created/modified in the last $DaysToInvestigate days" -ForegroundColor Yellow
        $InboundConnectors|Export-Csv -NoTypeInformation -Path "$ExportPath\InboundConnectors.csv"
    }
    else{Write-Host "No Inbound OnPrem Connectors created/modified in the past $DaysToInvestigate days" -ForegroundColor Green}
	
	if($null -ne $OutboundConnectors)
    {
        Write-Host "Outbound Connectors have been created/modified in the last $DaystoInvestigate days" -ForegroundColor Yellow
        $OutboundConnectors|Export-Csv -NoTypeInformation -Path "$ExportPath\OutboundConnectors.csv"
    }
    else{Write-Host "No Outbound Connectors created/modified in the past $DaysToInvestigate days" -ForegroundColor Green}
    
    return $InboundConnectors, $OutboundConnectors
}



<#
Transport Rules - Done
Forwarding
Redirect - Done
Journaling - Done
CBR - Done
BCC - Done
Audit 14
#>


Function Get-SuspiciousTransportRules
{   
    #ToDo - only notify about recent rules
    $SuspiciousTransportRules = @()
    $TransportRules = Get-TransportRule -ResultSize unlimited
    
    foreach($TransportRule in $TransportRules)
    {   
        switch -wildcard($TransportRule.Description)
        {   
            "*redirect the message to*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Route the message using the connector*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Blind carbon copy(Bcc) the message*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Forward the message*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Add the sender's manager as recipient type*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Send the incident report to*" 
                {$SuspiciousTransportRules += $TransportRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            default{}
        }
    }
    $SuspiciousTransportRules|Export-Csv -NoTypeInformation -Path "$ExportPath\TransportRulesToReview.csv"
    return  $SuspiciousTransportRules
    <#$AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
    -CmdletsToSearch "New-TransportRule","Set-TransportRule","Remove-TransportRule","Disable-TransportRule","Enable-TransportRule"#>
}

Function Get-SuspiciousJournalRule
{   
    #ToDo - only notify about recent Journal Rules
    $JournalRules = Get-JournalRule
    if($JournalRules.count -eq 0)
    {
        Write-Host "No Journal Rule" -ForegroundColor Green
        return $null
    }
    else 
    {
        Write-Host "We have detected the following Journal Rules:" -ForegroundColor Yellow
        $JournalRules|Format-Table Identity, Enabled, Scope, JournalEmailAddress, Recipient, WhenChanged
        $JournalRules|Export-Csv -NoTypeInformation -Path "$ExportPath\JournalRules.csv"
        return $JournalRules
    }
}


Function Get-GlobalAdminList
{
    $GlobalAdmins = Get-MsolRole | ForEach-Object{if (($_.name -eq "Company Administrator") -or ($_.name -eq "Exchange Service Administrator")) {$_}} |ForEach-Object{Get-MsolRoleMember -MaxResults 10000 -RoleObjectId $_.ObjectID}
    $GlobalAdminList = @()
    foreach($GlobalAdmin in $GlobalAdmins)
    {

        $MsolUser = get-msoluser -UserPrincipalName $GlobalAdmin.EmailAddress |Select-Object LastPasswordChangeTimestamp, StrongPasswordRequired
        $mailbox = get-mailbox $GlobalAdmin.EmailAddress -ErrorAction SilentlyContinue |Select-Object ForwardingAddress,ForwardingSmtpAddress, DeliverToMailboxAndForward
        
        $Admin = New-Object -TypeName psobject 
        $Admin | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $GlobalAdmin.EmailAddress 
        $Admin | Add-Member -MemberType NoteProperty -Name "LastPasswordChangeTimestamp" -Value $MsolUser.LastPasswordChangeTimestamp
        $Admin | Add-Member -MemberType NoteProperty -Name "MfaState" -Value $GlobalAdmin.StrongAuthenticationRequirements.State
        $Admin | Add-Member -MemberType NoteProperty -Name "StrongPasswordRequired" -Value $MsolUser.StrongPasswordRequired
        $Admin | Add-Member -MemberType NoteProperty -Name "ForwardingAddress" -Value $mailbox.ForwardingAddress
        $Admin | Add-Member -MemberType NoteProperty -Name "ForwardingSmtpAddress" -Value $mailbox.ForwardingSmtpAddress
        $Admin | Add-Member -MemberType NoteProperty -Name "DeliverToMailboxAndForward" -Value $mailbox.DeliverToMailboxAndForward

        $GlobalAdminList += $Admin
    }
    
    return $GlobalAdminList
}

Function Test-ProvisionedMailbox
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)

    [int]$i = 0
    while($i -lt $EmailAddresses.Count)
    {
        
        try     
        {
            $i++
            $GAExoMailbox = Get-EXOMailbox $EmailAddresses[$i-1] -ErrorAction Stop
            [string[]]$ProvisionedMailboxSMTPs += $GAExoMailbox.PrimarySmtpAddress
        }
        catch   {continue}
    }
    return [string[]]$ProvisionedMailboxSMTPs
}

Function Get-SuspiciousInboxRules
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)
    foreach($EmailAddress in $EmailAddresses)
    {
        $InboxRules += Get-InboxRule -Mailbox $EmailAddress
        Start-Sleep -Seconds 0.5
    }
    #ToDo - check which Admins have Inbox Rules with Forward/Redirect
    $InboxRules | Export-Csv -NoTypeInformation -Path "$ExportPath\GAInboxRules.csv"
    return $InboxRules
}

#region GA audit disable & audit bypass
Function Get-EXOAuditBypass
{
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
}
#endregion GA audit disable & audit bypass

function Get-CompromisedAdminAudit
{
    #Loading Dependencies from other APs
    . $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1

    #Individual AdminAudit Calls for suspicious actions
    $InboxRuleAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
                                    -CmdletsToSearch "New-InboxRule","Set-InboxRule","Remove-InboxRule","Enable-InboxRule","Disable-InboxRule"
    if($null -ne $InboxRuleAdminAudit)
    {$InboxRuleAdminAudit|Export-Csv -Append -NoTypeInformation -Path "$ExportPath\EXOAdminAuditLogs.csv"}

    $InboundConnectorAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
                                                -CmdletsToSearch "New-InboundConnector","Set-InboundConnector", "Remove-InboundConnector"
    if($null -ne $InboundConnectorAdminAudit)
    {$InboundConnectorAdminAudit|Export-Csv -Append -NoTypeInformation -Path "$ExportPath\EXOAdminAuditLogs.csv"}

    $OutboundConnectorAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
                                                -CmdletsToSearch "New-OutboundConnector","Set-OutboundConnector", "Remove-OutboundConnector"
    if($null -ne $OutboundConnectorAdminAudit)
    {$OutboundConnectorAdminAudit|Export-Csv -Append -NoTypeInformation -Path "$ExportPath\EXOAdminAuditLogs.csv"}

    $TransportRuleAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
    -CmdletsToSearch "New-TransportRule","Set-TransportRule","Remove-TransportRule","Disable-TransportRule","Enable-TransportRule"
    if($null -ne $OutboundConnectorAdminAudit)
    {$TransportRuleAdminAudit|Export-Csv -Append -NoTypeInformation -Path "$ExportPath\EXOAdminAuditLogs.csv"}

    return $InboundConnectorAdminAudit,$OutboundConnectorAdminAudit,$TransportRuleAdminAudit,$InboxRuleAdminAudit
}

Function Export-CompromisedHTMLReport
{param(
    [Object[]][Parameter(Mandatory=$false)] $InboxRules,
    [Object[]][Parameter(Mandatory=$false)] $TransportRules,
    [Object[]][Parameter(Mandatory=$false)] $InboundConnectors,
    [Object[]][Parameter(Mandatory=$false)] $OutboundConnectors,
    [Object[]][Parameter(Mandatory=$false)] $JournalRules
)
    #Export Info to HTML
    $header = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    .ResultOk {
        color: #008000;
    }
    
    .ResultNotOk {
        color: #ff0000;
    }
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

        #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }
    



</style>
"@
    #ToDo - validate if any Param null -> PreContent with Message that we have no Suspicious Entries for that.
    $ReportTitle = "<h1>Compromised Investigation</h1>"

    if($TransportRules)
    {
        $TransportRules = $SuspiciousTransportRules | ConvertTo-Html -Property Name,Description,State,Guid,WhenChanged -Fragment `
                                        -PreContent "<h2 class=`"ResultNotOk`">Suspicious Transport Rules</h2>"
    }
    else 
    {
        $TransportRules = $TransportRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Transport Rules</h2>"
    }

    if($InboxRules)
    {
        $InboxRules = $GAInboxRules | ConvertTo-Html -Property Name,Description,Enabled -Fragment -PreContent "<h2>Suspicious Inbox Rules</h2>"
    }
    else 
    {
        $InboxRules = $InboxRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Inbox Rules</h2>"
    }

    if($InboundConnectors)
    {
        $InboundConnectors = $InboundConnectors | ConvertTo-Html -Property Name, Enabled, WhenChangedUTC -As List -PreContent "<h2>Suspicious Inbound Connectors</h2>"
    }
    else 
    {
        $InboundConnectors = $InboundConnectors | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Inbound Connectors</h2>"
    }

    if($OutboundConnectors)
    {
        $OutboundConnectors = $OutboundConnectors | ConvertTo-Html -Property Name, Enabled, WhenChangedUTC -As List -PreContent "<h2>Suspicious Outbound Connectors</h2>"
    }
    else 
    {
        $OutboundConnectors = $OutboundConnectors | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Outbound Connectors</h2>"
    }

    if($JournalRules)
    {
        $JournalRules = $JournalRules | ConvertTo-Html -Property Name, Enabled, WhenChangedUTC -As List -PreContent "<h2>Suspicious Journal Rules</h2>"
    }
    else 
    {
        $JournalRules = $JournalRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Rules</h2>"
    }
    
    $Report = ConvertTo-Html -Head $header -Body "$ReportTitle $InboxRules $TransportRules $InboundConnectors $OutboundConnectors $JournalRules" `
                                -Title "Compromised Investigation" -PreContent "<p>Creation Date: $now</p>"

    $Report | Out-File "$ExportPath\CompromisedReport_$ts.htm"
}
Function Start-CompromisedMain
{   
    Clear-Host

    #Connect to O365 Workloads
    $Workloads = "Exo2", "MSOL"#, "AAD", "SCC"
    
    Connect-O365PS $Workloads

    $CurrentProperty = "Connecting to: $Workloads"
    
    $CurrentDescription = "Success"
    
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

    #Create Log Path
    $ts= get-date -Format yyyyMMdd_HHmmss
    $ExportPath = "$global:WSPath\Compromised_$ts"
    mkdir $ExportPath -Force|Out-Null
    
    $now = (Get-date).ToUniversalTime() #([datetime]::UtcNow)
    
    $DaysToInvestigate = 14

    $GlobalAdminList = Get-GlobalAdminList
    $GlobalAdminList | Export-Csv -NoTypeInformation -Path "$ExportPath\GlobalAdminList.csv"

    [string[]]$GASMTPs = $GlobalAdminList.UserPrincipalName

    [string[]]$ProvisionedMailboxSMTPs = Test-ProvisionedMailbox -EmailAddresses $GASMTPs

    if($ProvisionedMailboxSMTPs.Count -gt 0)
    {   
        $GAInboxRules = Get-SuspiciousInboxRules -EmailAddresses $ProvisionedMailboxSMTPs
    }

    $InboundConnectors, $OutboundConnectors = Get-RecentSuspiciousConnectors -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now

    Get-EXOAuditBypass

    $JournalRules = Get-SuspiciousJournalRule

    $SuspiciousTransportRules = Get-SuspiciousTransportRules  
    
    $BlockSenderReasons = Get-BlockedSenderReasons

    $InboundConnectorAdminAudit,$OutboundConnectorAdminAudit,$TransportRuleAdminAudit,$InboxRuleAdminAudit = Get-CompromisedAdminAudit

    #Call Azure AD Sign In and collect login audit for admins

    Export-CompromisedHTMLReport -InboundConnectors $InboundConnectors -OutboundConnectors $OutboundConnectors `
                        -InboxRules $GAInboxRules -TransportRules $SuspiciousTransportRules
    
    Write-Host "Exported logs to $ExportPath, you will be returned to O365Troubleshooters Main Menu" -ForegroundColor Green
    
    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}
