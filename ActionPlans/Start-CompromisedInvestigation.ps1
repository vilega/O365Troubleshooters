Function Get-BlockedSenderReasons
{
    ###Get Blocked Senders and Create Hashtable Array with SenderAddress & Reasons
    $blockedSenders = Get-BlockedSenderAddress

    if($null -ne $blockedSenders)
    {
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
        return "No Banned Outbound Senders Found"
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

    $AdminAuditLogs = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate `
    -CmdletsToSearch "New-TransportRule","Set-TransportRule","Remove-TransportRule","Disable-TransportRule","Enable-TransportRule"
}

Function Get-SuspiciousJournalRule
{
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
    
    <#foreach($EmailAddress in $EmailAddresses)
    {   
        Write-Host "Email Address #$i"
        $i++
        try     
        {
            $GAExoMailbox = Get-EXOMailbox $EmailAddress -ErrorAction Stop
            $ProvisionedMailboxSMTPs += $GAExoMailbox.PrimarySmtpAddress
        }
        catch   {continue}        
    }#>

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

Function Start-CompromisedMain
{   
    Clear-Host
    #Loading Dependencies from other APs
    . $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1

    #Connect to O365 Workloads
    $Workloads = "exo", "MSOL"#, "AAD", "SCC"
    
    Connect-O365PS $Workloads

    $CurrentProperty = "Connecting to: $Workloads"
    
    $CurrentDescription = "Success"
    
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

    #Create Log Path
    $ts= get-date -Format yyyyMMdd_HHmmss
    $ExportPath = "$global:WSPath\Compromised_$ts"
    mkdir $ExportPath -Force
    
    $now = (Get-date).ToUniversalTime() #([datetime]::UtcNow)
    
    $DaysToInvestigate = 14

    $GlobalAdminList = Get-GlobalAdminList
    $GlobalAdminList | Export-Csv -NoTypeInformation -Path "$ExportPath\GlobalAdminList.csv"

    [string[]]$GASMTPs = $GlobalAdminList.UserPrincipalName

    $ProvisionedMailboxSMTPs = Test-ProvisionedMailbox -EmailAddresses $GASMTPs
    
    if($ProvisionedMailboxSMTPs.Count -gt 0)
    {
        $GAInboxRules = Get-SuspiciousInboxRules -EmailAddresses $ProvisionedMailboxSMTPs
        $GAInboxRules | Export-Csv -NoTypeInformation -Path "$ExportPath\GAInboxRules.csv"
    }

    $InboundConnectors, $OutboundConnectors = Get-RecentSuspiciousConnectors -DaysToInvestigate $DaysToInvestigate -CurrentDateTime $now
    
    <#$InboundConnectorAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate -CmdletsToSearch "New-InboundConnector","Set-InboundConnector", `
                                                                                                        "Remove-InboundConnector"
    $InboundConnectorAdminAudit|Export-Csv -NoTypeInformation -Path "$ExportPath\InboundConnectorAdminAudit.csv"#>

    <#$OutboundConnectorAdminAudit = Search-EXOAdminAudit -DaysToSearch $DaysToInvestigate -CmdletsToSearch "New-OutboundConnector","Set-OutboundConnector", `
                                                                                                            "Remove-OutboundConnector"
    $OutboundConnectorAdminAudit|Export-Csv -NoTypeInformation -Path "$ExportPath\OutboundConnectorAdminAudit.csv"#>


    Get-EXOAuditBypass

    Write-Host "Exported logs to $ExportPath, you will be returned to O365Troubleshooters Main Menu" -ForegroundColor Green
    
    Read-Key

    Clear-Host

    Start-O365TroubleshootersMenu
}
