Function Get-BlockedSenderReasons
{param([bool][Parameter(Mandatory=$true)] $isFormatted)
    $blockedSenders = Get-BlockedSenderAddress
    if($null -ne $blockedSenders)
    {
        if($isFormatted)
        {  
            $blockedSenders|Export-Csv -NoTypeInformation -Path "$ExportPath\BlockedOutboundSenders.csv"
            
            $blockedSenderReasons = @()

            foreach($blockedSender in $blockedSenders)
            {
                $Reason = $blockedSender.Reason.Replace(";","`n")
                $Reason = ConvertFrom-StringData $Reason.Replace(":","=")
                $Reason["SenderAddress"] = $blockedSender.SenderAddress
                [hashtable[]]$blockedSenderReasons += $Reason
            }
        }
        else   
        {
            $blockedSenderReasons = $blockedSenders
        }
        return $BlockedSenderReasons
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
        Write-Warning "Inbound On Premises connectors have been created/modified in the last $DaysToInvestigate days"
        $InboundConnectors|Export-Csv -NoTypeInformation -Path "$ExportPath\InboundConnectors.csv"
    }
    else{Write-Host "No Inbound OnPrem Connectors created/modified in the past $DaysToInvestigate days" -ForegroundColor Green}
	
	if($null -ne $OutboundConnectors)
    {
        Write-Warning "Outbound Connectors have been created/modified in the last $DaystoInvestigate days"
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
    $SuspiciousTransportRules = @()
    $TransportRules = Get-TransportRule -ResultSize unlimited | Sort-Object -Descending WhenChanged
    
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
    $JournalRules = Get-JournalRule
    if($JournalRules.count -eq 0)
    {
        Write-Host "No Journal Rule" -ForegroundColor Green
        return $null
    }
    else 
    {
        Write-Warning "We have detected Journal Rules"
        #$JournalRules|Format-Table Identity, Enabled, Scope, JournalEmailAddress, Recipient, WhenChanged
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
            $GAExoMailbox = Get-Mailbox $EmailAddresses[$i-1] -ErrorAction Stop
            [string[]]$ProvisionedMailboxSMTPs += $GAExoMailbox.PrimarySmtpAddress
        }
        catch   {continue}
    }
    return [string[]]$ProvisionedMailboxSMTPs
}

Function Get-SuspiciousInboxRules
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)

    $InboxRules = @()
    $SuspiciousInboxRules = @()
    foreach($EmailAddress in $EmailAddresses)
    {
        $InboxRules += Get-InboxRule -Mailbox $EmailAddress | Select-Object *,@{Name="Mailbox";expression={$EmailAddress}}
        Start-Sleep -Seconds 0.5
    }

    foreach($InboxRule in $InboxRules)
    {   
        switch -wildcard($InboxRule.Description)
        {   
            "*redirect the message to*" 
                {$SuspiciousInboxRules += $InboxRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*move the message to*" 
                {$SuspiciousInboxRules += $InboxRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*Forward the message*" 
                {$SuspiciousInboxRules += $InboxRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            "*delete the message*" 
                {$SuspiciousInboxRules += $InboxRule<#|Select-Object Name,Description,State,Guid,WhenChanged#>}
            default{}
        }
    }
    #ToDo - check which Admins have Inbox Rules with Forward/Redirect/Delete/MoveToFolder
    $SuspiciousInboxRules | Export-Csv -NoTypeInformation -Path "$ExportPath\GAInboxRules.csv"
    return $SuspiciousInboxRules
}

#region GA audit disable & audit bypass
Function Get-EXOAuditBypass
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)

    if ((Get-OrganizationConfig).AuditDisabled -eq $true)
    {
        Write-Warning "Automatic AuditEnabled at organization level is turned off"
        return $true, $null, $null
    }

    else
    {
        Write-Host "Automatic AuditEnabled at organization level is turned on" -ForegroundColor Green
        
        foreach($GASMTP in $GASMTPs)
        {   
            $GAMailbox = Get-Mailbox $GASMTP -ErrorAction SilentlyContinue
            if (($GAMailbox).AuditEnabled -eq $false)
            {
                Write-Warning "The following Global Administrator $($GASMTP) has mailbox audit disabled"
                [Object[]]$MailboxAuditDisabledGAs += $GAMailbox
            }
            
            $GAMailboxAuditBypass = Get-MailboxAuditBypassAssociation -Identity $GASMTP
            if (($GAMailboxAuditBypass).AuditByPassEnabled -eq $true)
            {
                Write-Warning "The following administrator's ($GASMTP) actions on other mailboxes are not audited"
                [Object[]]$MailboxAuditBypassGAs += $GAMailboxAuditBypass
            }
        }
        return $false, $MailboxAuditDisabledGAs, $MailboxAuditBypassGAs
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

function Get-GAAzureSignInLogs
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)
    #Dot Sourcing Start-AzureADAuditSignInLogSearch.ps1
    . $script:modulePath\ActionPlans\Start-AzureADAuditSignInLogSearch.ps1

    if($DaysToInvestigate -gt 30)
    {
        $DaysToSearch = 30
        Write-Warning "The Compromised Investigation was scoped to $DaysToInvestigate days
For Global Admin Azure AD Sign In logs we will be able to provide a maximum of 30 days of logs"
    }
    else 
    {
        $DaysToSearch = $DaysToInvestigate
    }

    foreach($GASMTP in $EmailAddresses)
    {   
        Search-AzureAdSignInAudit -DaysToSearch $DaysToSearch -Upn $GASMTP
        $global:AzureAdSignInAll | Export-Csv "$ExportPath\GA_AllSignInAuditLogs_$ts.csv" -Append -NoTypeInformation
        $global:AzureAdSignInFail | Export-Csv "$ExportPath\GA_FailSignInAuditLogs_$ts.csv" -Append -NoTypeInformation
    }
}

function Get-GlobalAdminsWithIssues
{param([Object[]][Parameter(Mandatory=$true)] $GlobalAdminList)
    foreach($GlobalAdmin in $GlobalAdminList)
    {
        if( ($GlobalAdmin.MfaState -notmatch "Enforced") -or ($GlobalAdmin.StrongPasswordRequired -ne $true) `
                -or ($null -ne $GlobalAdmin.ForwardingAddress) -or ($null -ne $GlobalAdmin.ForwardingSmtpAddress))
        {
            [PSCustomObject[]]$GlobalAdminsWithIssues += $GlobalAdmin
        }
    }
    return $GlobalAdminsWithIssues
}
Function Export-CompromisedHTMLReport
{param(
    [Object[]][Parameter(Mandatory=$false)] $InboxRules,
    [Object[]][Parameter(Mandatory=$false)] $TransportRules,
    [Object[]][Parameter(Mandatory=$false)] $InboundConnectors,
    [Object[]][Parameter(Mandatory=$false)] $OutboundConnectors,
    [Object[]][Parameter(Mandatory=$false)] $JournalRules,
    [Object[]][Parameter(Mandatory=$false)] $GlobalAdminsWithIssues,
    [Object[]][Parameter(Mandatory=$false)] $BlockedSenderReasons,
    [Object[]][Parameter(Mandatory=$false)] $MailboxAuditDisabledGAs,
    [Object[]][Parameter(Mandatory=$false)] $MailboxAuditBypassGAs,
    [bool][Parameter(Mandatory=$false)] $OrganizationMailboxAuditDisabled
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
<a href="https://www.powershellgallery.com/packages/O365Troubleshooters" target="_blank">
  <img src="$script:modulePath\Resources\O365Troubleshooters-Logo.png" alt="O365Troubleshooters" width="173" height="128">
</a>
<img src=>
"@

    $ReportTitle = "<h1>Compromised Investigation</h1>"

    if($TransportRules)
    {
        $TransportRules = $SuspiciousTransportRules | ConvertTo-Html -Property Name, Description, State, Guid, WhenChanged -Fragment `
                                        -PreContent "<h2 class=`"ResultNotOk`">Suspicious Transport Rules</h2>"
    }
    else 
    {
        $TransportRules = $TransportRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Transport Rules</h2>"
    }

    if($InboxRules)
    {
        $InboxRules = $GAInboxRules | ConvertTo-Html -Property Mailbox, Name, Description, Enabled -Fragment `
                                -PreContent "<h2 class=`"ResultNotOk`">Suspicious Global Admin Inbox Rules</h2>"
    }
    else 
    {
        $InboxRules = $InboxRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Global Admin Inbox Rules</h2>"
    }

    if($InboundConnectors)
    {
        $InboundConnectors = $InboundConnectors | ConvertTo-Html -Property Name, Enabled, ConnectorType, SenderIPAddresses, SenderDomains, `
                                                                    TlsSenderCertificateName, WhenChangedUTC -Fragment <#-As List#> `
                                                        -PreContent "<h2 class=`"ResultNotOk`">Suspicious Inbound Connectors</h2>"
    }
    else 
    {
        $InboundConnectors = $InboundConnectors | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Inbound Connectors</h2>"
    }

    if($OutboundConnectors)
    {
        $OutboundConnectors = $OutboundConnectors | ConvertTo-Html -Property Name, Enabled, ConnectorType, UseMXRecord, SmartHosts, `
                                                                    WhenChangedUTC -Fragment <#-As List#> `
                                                            -PreContent "<h2 class=`"ResultNotOk`">Suspicious Outbound Connectors</h2>"
    }
    else 
    {
        $OutboundConnectors = $OutboundConnectors | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Outbound Connectors</h2>"
    }

    if($JournalRules)
    {
        $JournalRules = $JournalRules | ConvertTo-Html -Property Name, Enabled, Scope, JournalEmailAddress -Fragment <#-As List#> `
                                -PreContent "<h2 class=`"ResultNotOk`">Suspicious Journal Rules</h2>"
    }
    else 
    {
        $JournalRules = $JournalRules | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Suspicious Journal Rules</h2>"
    }

    if($GlobalAdminsWithIssues)
    {
        $GlobalAdminsWithIssues = $GlobalAdminsWithIssues | ConvertTo-Html -PreContent "<h2 class=`"ResultNotOk`">Global Admins configuration issues</h2>"
    }
    else 
    {
        $GlobalAdminsWithIssues = $GlobalAdminsWithIssues | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Global Admin configuration security concerns</h2>"
    }

    if($BlockedSenderReasons)
    {
        $BlockedSenderReasons = $BlockedSenderReasons | ConvertTo-Html -Property SENDERADDRESS, REASON, CREATEDDATETIME -PreContent "<h2 class=`"ResultNotOk`">Restricted Users found</h2>"
    }
    else 
    {
        $BlockedSenderReasons = $BlockedSenderReasons | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Restricted Users found</h2>"
    }

    $HiddenInboxRulesWarningString = "This Script does not currently check for Hidden Inbox Rules.
To identify and delete such rules, perform the steps from the following article:
<a href=`"https://docs.microsoft.com/en-us/archive/blogs/hkong/how-to-delete-corrupted-hidden-inbox-rules-from-a-mailbox-using-mfcmapi`" target=`"_blank`">How To Check and Delete Corrupted or Hidden Inbox Rules</a>"
<#    $HiddenInboxRulesWarning = New-Object -TypeName psobject 
    $HiddenInboxRulesWarning | Add-Member -MemberType NoteProperty -Name "Warning" -Value $HiddenInboxRulesWarningString #>
    $HiddenInboxRulesWarning = $null | ConvertTo-Html -PostContent $HiddenInboxRulesWarningString -PreContent "<h2 class=`"ResultNotOk`">Hidden Inbox Rules</h2>"

    if(!$OrganizationMailboxAuditDisabled)
    {
        if($MailboxAuditDisabledGAs)
        {
            $MailboxAuditDisabledGAs = $MailboxAuditDisabledGAs | ConvertTo-Html -Property DisplayName, PrimarySmtpAddress, AuditEnabled `
                                                                    -PreContent "<h2 class=`"ResultNotOk`">Global Admin Mailboxes with Audit Disabled</h2>"
        }
        else 
        {
            $MailboxAuditDisabledGAs = $MailboxAuditDisabledGAs | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Global Admin Mailboxes have Audit Disabled</h2>" 
        }

        if($MailboxAuditBypassGAs)
        {
            $MailboxAuditBypassGAs = $MailboxAuditBypassGAs|Select-Object Identity,@{Name="PrimarySmtpAddress";expression={(Get-Recipient -Identity $_.DistinguishedName).PrimarySmtpAddress}}, AuditBypassEnabled | `
                                                            ConvertTo-Html -PreContent "<h2 class=`"ResultNotOk`">Global Admin with Audit Bypass on other Mailboxes</h2>"
        }
        else 
        {
            $MailboxAuditBypassGAs = $MailboxAuditBypassGAs | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">No Global Admin with Audit Bypass on other Mailboxes found</h2>" 
        }
        $OrganizationMailboxAuditDisabledWarning = $null | ConvertTo-Html -PreContent "<h2 class=`"ResultOk`">Mailbox Audit Enabled Organization Wide</h2>"
    }
    else 
    {
        $OrganizationMailboxAuditDisabledWarning = $null | ConvertTo-Html -PreContent "<h2 class=`"ResultNotOk`">Mailbox Audit Disabled Organization Wide, check via Get-OrganizationConfig|select AuditDisabled</h2>"
    }

    $AdminAuditNotificationString = "We have exported Full Admin Audit Logs for cmdlets:<br>
    &emsp;New-InboxRule, Set-InboxRule, Remove-InboxRule, Enable-InboxRule, Disable-InboxRule<br>
    &emsp;New-InboundConnector, Set-InboundConnector, Remove-InboundConnector<br>
    &emsp;New-OutboundConnector, Set-OutboundConnector, Remove-OutboundConnector<br>
    &emsp;New-TransportRule, Set-TransportRule, Remove-TransportRule, Disable-TransportRule, Enable-TransportRule<br>
These logs can be found in file:<br>
$ExportPath\EXOAdminAuditLogs.csv"

    $AdminAuditNotification = $null | ConvertTo-Html -PostContent $AdminAuditNotificationString -PreContent "<h2 class=`"ResultNotOk`">Admin Audit Logs</h2>"

    $GlobalAdminsSignInAuditLogsNotificationString = "We have exported the following sign-in logs for global admins:<br>
    &emsp;AllSignInAuditLogs_$ts.csv - contains all audit sign-in log for global admins<br>
    &emsp;FailSignInAuditLogs_$ts.csv - contains fail audit sign-in log for global admins<br>
    These logs can be found in file:<br>
    $ExportPath\EXOAdminAuditLogs.csv"
    $GlobalAdminsSignInAuditLogsNotification = $null | ConvertTo-Html -PostContent $GlobalAdminsSignInAuditLogsNotificationString -PreContent "<h2 class=`"ResultNotOk`">Admin Audit Sign-in Logs</h2>"
    
    $Report = ConvertTo-Html -Head $header -Body "$ReportTitle $GlobalAdminsWithIssues $HiddenInboxRulesWarning $InboxRules $OrganizationMailboxAuditDisabledWarning `
                                $MailboxAuditBypassGAs $MailboxAuditDisabledGAs $AdminAuditNotification $GlobalAdminsSignInAuditLogsNotification $BlockedSenderReasons $TransportRules $InboundConnectors $OutboundConnectors $JournalRules" `
                                -Title "Compromised Investigation" -PreContent "<p>Creation Date: $now</p>"

    $Report | Out-File "$ExportPath\CompromisedReport_$ts.htm"
}
Function Start-CompromisedMain
{   
    Clear-Host

    #Connect to O365 Workloads
    $Workloads = "Exo", "MSOL", "AzureADPreview"#, "AAD", "SCC"
    
    Connect-O365PS $Workloads

    $CurrentProperty = "Connecting to: $Workloads"
    
    $CurrentDescription = "Success"
    
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

    #Create Log Path
    $ts= get-date -Format yyyyMMdd_HHmmss
    $ExportPath = "$global:WSPath\Compromised_$ts"
    mkdir $ExportPath -Force|Out-Null
    
    $now = (Get-date).ToUniversalTime() #([datetime]::UtcNow)
    
    $DaysToInvestigate = Read-IntFromConsole -IntType "Number of days to investigate Tenant Compromise"

    $GlobalAdminList = Get-GlobalAdminList
    $GlobalAdminList | Export-Csv -NoTypeInformation -Path "$ExportPath\GlobalAdminList.csv"

    [string[]]$GASMTPs = $GlobalAdminList.UserPrincipalName

    [string[]]$ProvisionedMailboxSMTPs = Test-ProvisionedMailbox -EmailAddresses $GASMTPs

    if($ProvisionedMailboxSMTPs.Count -gt 0)
    {   
        $GAInboxRules = Get-SuspiciousInboxRules -EmailAddresses $ProvisionedMailboxSMTPs
    }

    $InboundConnectors, $OutboundConnectors = Get-RecentSuspiciousConnectors -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now

    $JournalRules = Get-SuspiciousJournalRule

    $SuspiciousTransportRules = Get-SuspiciousTransportRules  
    
    $BlockedSenderReasonsObject = Get-BlockedSenderReasons -isFormatted $false

    $InboundConnectorAdminAudit,$OutboundConnectorAdminAudit,$TransportRuleAdminAudit,$InboxRuleAdminAudit = Get-CompromisedAdminAudit

    Get-GAAzureSignInLogs -EmailAddresses $GASMTPs

    $GlobalAdminsWithIssues = Get-GlobalAdminsWithIssues -GlobalAdminList $GlobalAdminList

    $OrganizationMailboxAuditDisabled, $MailboxAuditDisabledGAs, $MailboxAuditBypassGAs = Get-EXOAuditBypass -EmailAddresses $GASMTPs

    Export-CompromisedHTMLReport -InboundConnectors $InboundConnectors -OutboundConnectors $OutboundConnectors `
                        -InboxRules $GAInboxRules -TransportRules $SuspiciousTransportRules -GlobalAdminsWithIssues $GlobalAdminsWithIssues `
                        -JournalRules $JournalRules -BlockedSenderReasons $BlockedSenderReasonsObject `
                        -OrganizationMailboxAuditDisabled $OrganizationMailboxAuditDisabled -MailboxAuditDisabledGAs $MailboxAuditDisabledGAs `
                        -MailboxAuditBypassGAs $MailboxAuditBypassGAs
    
    Write-Host -ForegroundColor Green "Exported logs to $ExportPath, here you will find:
    -HTML Summary Report
    $ExportPath\CompromisedReport_$ts.htm
    -Full CSV Output dump used for analysis and building HTML Report.
you will be returned to O365Troubleshooters Main Menu" 
    
    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}
