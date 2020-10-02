Function Get-BlockedSenderReasons {
[CmdletBinding()]
param([bool][Parameter(Mandatory=$true)] $isFormatted)

    [System.Collections.ArrayList]$BlockedSenders = @(Get-BlockedSenderAddress|Select-Object SenderAddress,Reason,CreatedDatetime)
    if($BlockedSenders)
    {   
        Write-Warning -Message "Found Senders Blocked due to Outbound Spam"
        if($isFormatted) {            
            $blockedSenderReasons = @()

            foreach($blockedSender in $blockedSenders) {
                $Reason = $blockedSender.Reason.Replace(";","`n")
                $Reason = ConvertFrom-StringData $Reason.Replace(":","=")
                $Reason["SenderAddress"] = $blockedSender.SenderAddress
                [hashtable[]]$blockedSenderReasons += $Reason
            }
        }
        else {
            $BlockedSenderReasons = $blockedSenders
        }
        $blockedSenders|Export-Csv -NoTypeInformation -Path "$ExportPath\BlockedOutboundSenders.csv"
        return $BlockedSenderReasons
    }
    else 
    {
        Write-Host -ForegroundColor Green "No Banned Outbound Senders Found"
        return $null
    }
}


Function Get-RecentSuspiciousConnectors
{
param(
    [int][Parameter(Mandatory=$true)] $DaysToInvestigate, 
    [datetime][Parameter(Mandatory=$true)] $CurrentDateTime
)
    [System.Collections.ArrayList]$SuspiciousInboundConnectors = @()
    [System.Collections.ArrayList]$SuspiciousOutboundConnectors = @()

    $InboundConnectors = @(Get-InboundConnector | Select-Object Name, Enabled, ConnectorType, SenderIPAddresses, SenderDomains, TlsSenderCertificateName, WhenCreatedUTC, WhenChangedUTC)
    foreach($InboundConnector in $InboundConnectors) {
        if( ($InboundConnector.ConnectorType -EQ "OnPremises") -and ( ($InboundConnector.WhenCreatedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) -or ($InboundConnector.WhenChangedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)))) 
        {
            $SuspiciousInboundConnectors.Add($InboundConnector)|Out-Null
        }
    }

    $OutboundConnectors = @(Get-OutboundConnector | Select-Object Name, Enabled, ConnectorType, UseMXRecord, SmartHosts, WhenCreatedUTC, WhenChangedUTC)
    foreach($OutboundConnector in $OutboundConnectors){
        if ( ($OutboundConnector.WhenCreatedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) -or ($OutboundConnector.WhenChangedUTC -ge $CurrentDateTime.AddDays(-$DaysToInvestigate)) )
        {
            $SuspiciousOutboundConnectors.Add($OutboundConnector)|Out-Null
        }
    }
    
    if($SuspiciousInboundConnectors) {
        Write-Warning "Inbound On Premises connectors have been created/modified in the last $DaysToInvestigate days"
        $InboundConnectors|Export-Csv -NoTypeInformation -Path "$ExportPath\SuspiciousInboundConnectors.csv"
    }
    else {
        Write-Host "No Inbound OnPrem Connectors created/modified in the past $DaysToInvestigate days" -ForegroundColor Green
    }
	
	if($SuspiciousOutboundConnectors) {
        Write-Warning "Outbound Connectors have been created/modified in the last $DaystoInvestigate days"
        $OutboundConnectors|Export-Csv -NoTypeInformation -Path "$ExportPath\SuspiciousOutboundConnectors.csv"
    }
    else {
        Write-Host "No Outbound Connectors created/modified in the past $DaysToInvestigate days" -ForegroundColor Green
    }
    
    return $SuspiciousInboundConnectors, $SuspiciousOutboundConnectors
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


Function Get-SuspiciousTransportRules {   
[CmdletBinding()]
param(
    [int][Parameter(Mandatory=$true)] $DaysToInvestigate, 
    [datetime][Parameter(Mandatory=$true)] $CurrentDateTime
)
    [System.Collections.ArrayList]$SuspiciousTransportRules = @()
    $TransportRules = @(Get-TransportRule -ResultSize unlimited | Select-Object Name,Description,State,Guid,WhenChanged)
    
    foreach($TransportRule in $TransportRules)
    {   
        if( $TransportRule.WhenChanged -ge $CurrentDateTime.AddDays(-$DaysToInvestigate) ) {
            switch -wildcard($TransportRule.Description)
            {   
                "*redirect the message to*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                "*Route the message using the connector*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                "*Blind carbon copy(Bcc) the message*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                "*Forward the message*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                "*Add the sender's manager as recipient type*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                "*Send the incident report to*" 
                    {$SuspiciousTransportRules.Add($TransportRule)}
                default{}
            }
        }

    }

    if($SuspiciousTransportRules) {
        Write-Warning -Message "Suspicious Transport Rules were found."
        $SuspiciousTransportRules|Export-Csv -NoTypeInformation -Path "$ExportPath\SuspiciousTransportRules.csv"
        return  $SuspiciousTransportRules
    }
    else {
        Write-Host "No Suspicious Transport Rules found" -ForegroundColor "Green"
        return  $null
    }
}

Function Get-SuspiciousJournalRule
{
[CmdletBinding()]
param(
    [int][Parameter(Mandatory=$true)] $DaysToInvestigate, 
    [datetime][Parameter(Mandatory=$true)] $CurrentDateTime
)    
    [System.Collections.ArrayList]$SuspiciousJournalRules = @()
    $JournalRules = @(Get-JournalRule|Select-Object Identity,Guid,JournalEmailAddress,Recipient,Scope,Enabled,WhenChanged)
 
    foreach($JournalRule in $JournalRules) {
        if( $JournalRule.WhenChanged -ge ($CurrentDateTime.AddDays(-$DaysToInvestigate)) ){
            $SuspiciousJournalRules.Add($JournalRule)|Out-Null
        }
    }

    if(!$SuspiciousJournalRules) {
        Write-Host "No Journal Rule" -ForegroundColor Green
        return $null
    }
    else{
        Write-Warning "We have detected Journal Rules"
        #$JournalRules|Format-Table Identity, Enabled, Scope, JournalEmailAddress, Recipient, WhenChanged
        $SuspiciousJournalRules|Export-Csv -NoTypeInformation -Path "$ExportPath\JournalRules.csv"
        return $SuspiciousJournalRules
    }
}


Function Get-GlobalAdminList
{
    $GlobalAdmins = Get-MsolRole | ForEach-Object{if (($_.name -eq "Company Administrator") -or ($_.name -eq "Exchange Service Administrator")) {$_}} |ForEach-Object{Get-MsolRoleMember -MaxResults 10000 -RoleObjectId $_.ObjectID}
    [System.Collections.ArrayList]$GlobalAdminList = @()
    foreach($GlobalAdmin in $GlobalAdmins){
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

        $GlobalAdminList.Add($Admin)|Out-Null
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
{
param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)

    [System.Collections.ArrayList]$SuspiciousInboxRules = @()
    foreach($EmailAddress in $EmailAddresses)
    {
        [System.Collections.ArrayList]$InboxRules = @(Get-InboxRule -Mailbox $EmailAddress|Select-Object @{Name="Mailbox";expression={$EmailAddress}},Name, Description, Identity, Enabled)
        
        foreach($InboxRule in $InboxRules)
        {   
            switch -wildcard($InboxRule.Description)
            {   
                "*redirect the message to*" 
                {
                    $SuspiciousInboxRules.Add($InboxRule)|Out-Null
                }
                "*move the message to*" 
                {
                    $SuspiciousInboxRules.Add($InboxRule)|Out-Null
                }
                "*Forward the message*" 
                {
                    $SuspiciousInboxRules.Add($InboxRule)|Out-Null
                }
                "*delete the message*" 
                {
                    $SuspiciousInboxRules.Add($InboxRule)|Out-Null
                }
                default{Out-Null}
            }
        }
    }

    $SuspiciousInboxRules | Export-Csv -NoTypeInformation -Path "$ExportPath\GAInboxRules.csv"
    return $SuspiciousInboxRules
}

#region GA audit disable & audit bypass
Function Get-EXOAuditBypass
{param([string[]][Parameter(Mandatory=$true)] $EmailAddresses)
    [System.Collections.ArrayList]$MailboxAuditDisabledGAs = @()
    [System.Collections.ArrayList]$MailboxAuditBypassGAs = @()
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
            $GAMailbox = Get-Mailbox $GASMTP -ErrorAction SilentlyContinue|Select-Object PrimarySmtpAddress, AuditEnabled, WhenChangedUTC
            if (($GAMailbox).AuditEnabled -eq $false)
            {
                Write-Warning "The following Global Administrator $($GASMTP) has mailbox audit disabled"
                $MailboxAuditDisabledGAs.Add($GAMailbox)|Out-Null
            }
            
            $MailboxAuditBypassGA = Get-MailboxAuditBypassAssociation -Identity $GASMTP `
                                        | Select-Object @{Name="Mailbox";expression={$GASMTP}},Identity, AuditBypassEnabled, WhenChangedUTC
            if (($MailboxAuditBypassGA).AuditByPassEnabled -eq $true)
            {
                Write-Warning "The following administrator's ($GASMTP) actions on other mailboxes are not audited"
                $MailboxAuditBypassGAs.Add($MailboxAuditBypassGA)|Out-Null
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
{param([System.Collections.ArrayList][Parameter(Mandatory=$true)] $GlobalAdminList)
    [System.Collections.ArrayList]$GlobalAdminsWithIssues = @()
    foreach($GlobalAdmin in $GlobalAdminList)
    {
        if( ($GlobalAdmin.MfaState -notmatch "Enforced") -or ($GlobalAdmin.StrongPasswordRequired -ne $true) `
                -or ($null -ne $GlobalAdmin.ForwardingAddress) -or ($null -ne $GlobalAdmin.ForwardingSmtpAddress))
        {
            $GlobalAdminsWithIssues.Add($GlobalAdmin)
        }
    }
    return $GlobalAdminsWithIssues
}
Function Export-CompromisedHTMLReport
{param(
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $InboxRules,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $TransportRules,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $InboundConnectors,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $OutboundConnectors,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $JournalRules,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $GlobalAdminsWithIssues,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $BlockedSenderReasons,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $MailboxAuditDisabledGAs,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)] $MailboxAuditBypassGAs,
    [String][Parameter(Mandatory=$false)]$HTMLFilePath,
    [bool][Parameter(Mandatory=$false)] $OrganizationMailboxAuditDisabled
)
    #Export Info to HTML
    <#$header = @"
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
    $ReportTitle = "<h1>Compromised Investigation</h1>"#>
    #Prepare requirements for Module HTML Functions
    $TableType = "Table"
    [string]$NoIssue = "We have not identified any configuration issues."
    [System.Collections.ArrayList]$Office365RelayHTMLReportArray = @()
    
    #Global Admins with Issues
    [string]$SectionTitle = "Global Admin with Issues"
    [string]$Description = "This section should display configuration snapshots for Admins where we have found potential issues."
    if($GlobalAdminsWithIssues) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $GlobalAdminsWithIssues -TableType $TableType

        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Inbox Rules
    [string]$SectionTitle = "Global Admin Suspicious Inbox Rules"
    [string]$Description = "This section should display suspicious Inbox Rules we have identified on GA Mailboxes.<br>The Script does yet check for Hidden Inbox Rules, you can do this manually as show in article:<br><a href=`"https://docs.microsoft.com/en-us/archive/blogs/hkong/how-to-delete-corrupted-hidden-inbox-rules-from-a-mailbox-using-mfcmapi`" target=`"_blank`">How To Check and Delete Corrupted or Hidden Inbox Rules</a>"
    if($InboxRules) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "CustomObject" -EffectiveDataArrayList $InboxRules -TableType $TableType
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else {
        $SectionTitleColor = "Black"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Hidden Inbox Rules
    <#$HiddenInboxRulesWarningString = "This Script does not currently check for Hidden Inbox Rules.<br>
To identify and delete such rules, perform the steps from the following article:<br>
<a href=`"https://docs.microsoft.com/en-us/archive/blogs/hkong/how-to-delete-corrupted-hidden-inbox-rules-from-a-mailbox-using-mfcmapi`" target=`"_blank`">How To Check and Delete Corrupted or Hidden Inbox Rules</a>"#>
    <#    $HiddenInboxRulesWarning = New-Object -TypeName psobject 
    $HiddenInboxRulesWarning | Add-Member -MemberType NoteProperty -Name "Warning" -Value $HiddenInboxRulesWarningString #>
    #$HiddenInboxRulesWarning = $null | ConvertTo-Html -PostContent $HiddenInboxRulesWarningString -PreContent "<h2 class=`"ResultNotOk`">Hidden Inbox Rules</h2>"
    <#
    $SectionTitleColor = "Red"
    $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                    -DataType "String" -EffectiveDataString $HiddenInboxRulesWarningString
    $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    #>

    #Journal Rules
    [string]$SectionTitle = "Suspicious Journal Rules"
    [string]$Description = "This section should display Suspicious Journal Rules that were identified."
    if($JournalRules)
    {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "CustomObject" -EffectiveDataArrayList $JournalRules -TableType $TableType
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else 
    {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Blocked Senders
    [string]$SectionTitle = "Blocked Senders - Outbound Spam"
    [string]$Description = "This section should display Blocked Senders from your Organization, identified as Outbound Spam Senders."
    if($BlockedSenderReasons) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "CustomObject" -EffectiveDataArrayList $BlockedSenderReasons -TableType $TableType
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Inbound Connectors
    [string]$SectionTitle = "Suspicious Inbound Connectors"
    [string]$Description = "This section should display Suspicious Inbound Connectors, which can be used by attackers to relay emails through your tenant."
    if($InboundConnectors) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "CustomObject" -EffectiveDataArrayList $InboundConnectors -TableType $TableType
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Outbound Connectors
    [string]$SectionTitle = "Suspicious Outbound Connectors"
    [string]$Description = "This section should display Suspicious Outbound Connectors, which can be used by attackers to route emails outside your tenant."
    if($OutboundConnectors) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "CustomObject" -EffectiveDataArrayList $OutboundConnectors -TableType $TableType
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }
    else {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Transport Rules
    [string]$SectionTitle = "Suspicious Transport Rules"
    [string]$Description = "This section should display Suspicious Transport Rules, which can be used by attackers to exfiltrate data from your organization."
    if($TransportRules) {
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $TransportRules -TableType $TableType

        $Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }
    else {
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Mailbox Audit checks
    [string]$SectionTitle = "Organization Wide Mailbox Audit"
    [string]$Description = "This section shows if Mailbox Auditing is Disabled Organization Wide."
    if(!$OrganizationMailboxAuditDisabled) {
        #Organization Wide Mailbox Audit
        $SectionTitleColor = "Green"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $NoIssue
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null

        #Global Admin Mailbox Audit Disabled
        [string]$SectionTitle = "Global Admin Mailboxes with Audit Disabled"
        [string]$Description = "This section shows Global Admin Mailboxes for which Mailbox Auditing is not enabled, this can be used by attackers to hide actions of exfiltrating data."
        if($MailboxAuditDisabledGAs) {
            $SectionTitleColor = "Red"
            $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                    -DataType "CustomObject" -EffectiveDataArrayList $MailboxAuditDisabledGAs -TableType $TableType
    
            $Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
        }
        else {
            $SectionTitleColor = "Green"
            $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                            -DataType "String" -EffectiveDataString $NoIssue
            $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
        }

        #Global Admins Bypassing Mailbox Auditing
        [string]$SectionTitle = "Global Admins that Bypass Mailbox Auditing"
        [string]$Description = "This section shows Global Admins that Bypass Mailbox Auditing, this can be used by attackers to hide actions of exfiltrating data."
        if($MailboxAuditBypassGAs) {
            $SectionTitleColor = "Red"
            $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                    -DataType "CustomObject" -EffectiveDataArrayList $MailboxAuditBypassGAs -TableType $TableType
    
            $Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
        }
        else {
            $SectionTitleColor = "Green"
            $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                            -DataType "String" -EffectiveDataString $NoIssue
            $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
        }
    }
    else {
        $OrgWideMailboxAuditDisabledWarning = "Mailbox Audit Disabled Organization Wide, check via Exchange Online Powershell : Get-OrganizationConfig|select AuditDisabled"
        $SectionTitleColor = "Red"
        $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                        -DataType "String" -EffectiveDataString $OrgWideMailboxAuditDisabledWarning
        $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null
    }

    #Admin Audit Log Notification
    [string]$SectionTitle = "Exchange Online Admin Audit Logs"
    [string]$Description = "This section provides information on Exchange Online Admin Audit logs for which raw data has been exported to CSV"
    $AdminAuditNotificationString = "We have exported Full Admin Audit Logs for cmdlets:<br>
&emsp;New-InboxRule, Set-InboxRule, Remove-InboxRule, Enable-InboxRule, Disable-InboxRule<br>
&emsp;New-InboundConnector, Set-InboundConnector, Remove-InboundConnector<br>
&emsp;New-OutboundConnector, Set-OutboundConnector, Remove-OutboundConnector<br>
&emsp;New-TransportRule, Set-TransportRule, Remove-TransportRule, Disable-TransportRule, Enable-TransportRule<br>
These logs can be found in file:<br>
$ExportPath\EXOAdminAuditLogs.csv"

    $SectionTitleColor = "Black"
    $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                    -DataType "String" -EffectiveDataString $AdminAuditNotificationString
    $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null

    #Global Admin Sign In Logs Notification
    [string]$SectionTitle = "Azure AD Global Admin Sign In Logs"
    [string]$Description = "This section provides information on Azure AD Global Admin Sign In logs for which raw data has been exporeted to CSV"
    $GlobalAdminsSignInAuditLogsNotificationString = "We have exported the following sign-in logs for global admins:<br>
&emsp;AllSignInAuditLogs_$ts.csv - contains all audit sign-in log for global admins<br>
&emsp;FailSignInAuditLogs_$ts.csv - contains fail audit sign-in log for global admins<br>
These logs can be found in file:<br>
$ExportPath\EXOAdminAuditLogs.csv"
    
    $SectionTitleColor = "Black"
    $HTMLReportEntry = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                                                    -DataType "String" -EffectiveDataString $GlobalAdminsSignInAuditLogsNotificationString
    $Office365RelayHTMLReportArray.Add($HTMLReportEntry) | Out-Null

    #Export HTML Report
    Export-ReportToHTML -FilePath $HTMLFilePath -PageTitle "Office 365 Compromised Tenant Investigation" `
                            -ReportTitle "Office 365 Compromised Tenant Investigation" `
                                -TheObjectToConvertToHTML $Office365RelayHTMLReportArray
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
    [string]$HTMLFilePath = "$ExportPath\CompromisedTenantInvestigaton.html"
    [System.Collections.ArrayList]$GAInboxRules = @()
    [System.Collections.ArrayList]$InboundConnectors = @()
    [System.Collections.ArrayList]$OutboundConnectors = @()
    [System.Collections.ArrayList]$JournalRules = @()
    [System.Collections.ArrayList]$SuspiciousTransportRules = @()
    [System.Collections.ArrayList]$BlockedSenderReasonsObject = @()
    [System.Collections.ArrayList]$GlobalAdminsWithIssues = @()
    [System.Collections.ArrayList]$MailboxAuditDisabledGAs = @()
    [System.Collections.ArrayList]$MailboxAuditBypassGAs = @()
    [System.Collections.ArrayList]$GlobalAdminList = Get-GlobalAdminList
    
    $GlobalAdminList | Export-Csv -NoTypeInformation -Path "$ExportPath\GlobalAdminList.csv"

    [string[]]$GASMTPs = $GlobalAdminList.UserPrincipalName

    [string[]]$ProvisionedMailboxSMTPs = Test-ProvisionedMailbox -EmailAddresses $GASMTPs

    if($ProvisionedMailboxSMTPs.Count -gt 0)
    {   
        $GAInboxRules = Get-SuspiciousInboxRules -EmailAddresses $ProvisionedMailboxSMTPs
    }

    $InboundConnectors, $OutboundConnectors = Get-RecentSuspiciousConnectors -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now

    [System.Collections.ArrayList]$JournalRules = @(Get-SuspiciousJournalRule -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now)

    [System.Collections.ArrayList]$SuspiciousTransportRules = @(Get-SuspiciousTransportRules -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now)
    
    [System.Collections.ArrayList]$BlockedSenderReasonsObject = @(Get-BlockedSenderReasons -isFormatted $false)

    $InboundConnectorAdminAudit,$OutboundConnectorAdminAudit,$TransportRuleAdminAudit,$InboxRuleAdminAudit = Get-CompromisedAdminAudit

    Get-GAAzureSignInLogs -EmailAddresses $GASMTPs

    $GlobalAdminsWithIssues = Get-GlobalAdminsWithIssues -GlobalAdminList $GlobalAdminList

    [System.Boolean]$OrganizationMailboxAuditDisabled, [System.Collections.ArrayList]$MailboxAuditDisabledGAs, [System.Collections.ArrayList]$MailboxAuditBypassGAs = Get-EXOAuditBypass -EmailAddresses $GASMTPs

    Export-CompromisedHTMLReport -InboundConnectors $InboundConnectors -OutboundConnectors $OutboundConnectors `
                        -InboxRules $GAInboxRules -TransportRules $SuspiciousTransportRules -GlobalAdminsWithIssues $GlobalAdminsWithIssues `
                        -JournalRules $JournalRules -BlockedSenderReasons $BlockedSenderReasonsObject `
                        -OrganizationMailboxAuditDisabled $OrganizationMailboxAuditDisabled -MailboxAuditDisabledGAs $MailboxAuditDisabledGAs `
                        -MailboxAuditBypassGAs $MailboxAuditBypassGAs -HTMLFilePath $HTMLFilePath
    
    Write-Host -ForegroundColor Green "Exported logs to $ExportPath, here you will find:
    -HTML Summary Report
    $ExportPath\CompromisedReport_$ts.htm
    -Full CSV Output dump used for analysis and building HTML Report.
you will be returned to O365Troubleshooters Main Menu" 
    
    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}
