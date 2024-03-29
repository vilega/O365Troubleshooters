#region Office365RelayDependencies
<#
Office 365 Relay Script Simulates a local application attempting to perform one of "SMTP Client Submission"/"SMTP Relay"/"Direct Send".
###leverages Send-MailMessage to handle the email submission.
###handles validation for Input data and error handling.
###has a central repository for known errors for this submission type and provides troubleshooting suggestions.

Aside from offering troubleshooting opportunities, it also grants option for customizing Email Body and Attachment.
Some built in features that can be used for AntiSpam testing:
###Standard test message body with timestamp
###Gtube test
###SpamLink SafeLink URL
###Input from Console
###*.htm / HTML formatted EML from Desktop
#>

<#
Validates Email Address via Regex
#>

# function Get-ValidEmailAddress([string]$EmailAddressType)
# {
#     [string]$EmailAddress = Read-Host "Enter Valid $EmailAddressType"

#     if($EmailAddress -match "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$")
#     {
#         return $EmailAddress
#     }
#     else
#     {
#         Get-ValidEmailAddress($EmailAddressType)
#     }
# }

<#
Validates Domain via Regex
Includes a Switch purely for future extention of the script
#>
function Get-ValidDomain([string]$DomainType)
{   
    switch($DomainType)
    {
        "Initial *.onmicrosoft.com"
        {
            [int]$count = 0
            do
            {
                Write-Host "Enter Valid $DomainType Domain Name: " -ForegroundColor Cyan -NoNewline
                [string]$Domain = Read-Host
                [bool]$valid = ($Domain.Trim() -match "^[A-Z0-9]+.onmicrosoft.com$")
                $count++
            }
            while (!$valid -and ($count -le 2))
            
            if ($valid)
            {
                return $Domain.Trim()
            }
            else 
            {
                Write-Host "3 invalid domain inputs received`r`n `
You will be returned to the O365Troubleshooters Main Menu!" -ForegroundColor Red
                Start-Sleep -Seconds 3
                Write-Log -function "Get-ValidDomain" -step "input address" -Description "After 3 invalid domain inputs, sending user back to O365Troubleshooters Main Menu"
                Exit-ScriptAndSaveLogs
            }
        }

        default
        {
        Write-Host "Unknown Domain Type Input Received"
        Get-RetryMenu
        }

    }
}


<#
Collects a List of Valid Recipient Email Addresses:
#Ask for Recipients
#Provide done to end list
#>

function Get-Office365RelayRecipients()
{   
    $Office365RelayErrorList.Clear()
    
    Write-Host -ForegroundColor Cyan "Enter a Number of Recipients : " -NoNewline

    [int]$Office365RelayRecipientCount = Read-Host -ErrorAction SilentlyContinue -ErrorVariable +Office365RelayErrorList
    
    if( ($null -eq $Office365RelayErrorList[0]) -and ($Office365RelayRecipientCount -gt 0 ) -and ($Office365RelayRecipientCount -le 500 ) )
    {   
        Write-Host -ForegroundColor Cyan "You will receive recurring prompts to enter each individual recipient"

        [string]$EmailAddressType = "RcptTo Email Address"

        [int]$i = 0
        
        while($i -lt $Office365RelayRecipientCount)
        {
            [string[]]$Office365RelayRecipients += Get-ValidEmailAddress([string]$EmailAddressType)
            $i++
        }

        return $Office365RelayRecipients
    }

    else 
    {
        Get-Office365RelayRecipients
    }
}

<#
Collect cloud smarthost for Smtp Relay/Direct Send function, performs:
    -Collects Cloud Smarthost
    -checks Resolve-DnsName to confirm that A record is correctly propagated for it.
#>
function Find-O365RelaySmarthost()
{
    $Office365RelayErrorList.Clear()

    [string]$DomainType = "Initial *.onmicrosoft.com"

    $InitialDomain = [string](Get-ValidDomain($DomainType))

    $MXResolution = Resolve-DnsName -Type MX -Name $InitialDomain -ErrorAction Continue -ErrorVariable +Office365RelayErrorList

    if ($MXResolution.Type -eq "MX")
    {   
        $Office365RelaySmarthost = Resolve-DnsName -Type A -Name $MXResolution.NameExchange -ErrorAction Continue `
                                    -ErrorVariable +Office365RelayErrorList

        if ($null -eq $Office365RelayErrorList[0])
        {
            return $Office365RelaySmarthost[0].Name
        }

        else
        {
            Get-ActionPlan("SmarthostFunctionErrors")
        }
    }

    else 
    {
        Get-ActionPlan("SmarthostFunctionErrors")
    }
}

<#
Collect Credentials function, performs :
    -clean-up for email address format based on Regex.
    -runs Powershell Connection check to validate credentials.
    -closes Powershell connection after confirming login works.
#>
function Get-AuthenticationCredentials()
{
    $Office365RelayErrorList.Clear()

    Write-Host -ForegroundColor Cyan "You will be prompted for Credentials of the account used to send emails via SMTP Client Submission"

    Read-Key

    $O365SenderCred = Get-Credential

    if ($null -eq $O365SenderCred)
    {
        Write-Host -ForegroundColor Red "UserName and/or Password empty`r`nYou will be returned to the Script Main Menu"
        Read-Key
        Get-MainMenu
    }

    $O365AuthenticationCredentialsSession = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365SenderCred `
        -Authentication Basic -AllowRedirection -ErrorAction Continue -ErrorVariable +Office365RelayErrorList
        
    if ($null -eq $Office365RelayErrorList[0])
    {
        Remove-PSSession -InstanceId $O365AuthenticationCredentialsSession.InstanceId
        return $O365SenderCred
    }

    else
    {
        Get-ActionPlan("AuthenticationFunctionErrors")
    }
}

<#
Test-Office365RelayScriptItemPath checks for item path
returns true or false (bool).
#>
function Test-Office365RelayScriptItemPath([string] $Office365RelayScriptFilePath)
{
    return Test-Path $Office365RelayScriptFilePath
}

<#
Get-MessageBody function provides:
-standard test body with timestamp
-Gtube test
-SpamLink SafeLink URL
-option to input from Console
-Custom content from htm files
#>
function Get-MessageBody([DateTime] $d)
{
    [string] $MessageBodyType = Read-Host "`r`nSelect a type of Message Body`r`n
A : Standard test message body with timestamp
B : Gtube test
C : SpamLink SafeLink URL
D : Input from Console
E : *.htm / HTML formatted EML from Desktop

Answer"

    switch($MessageBodyType.ToUpper())
    {
        A
        {
            return "This is the body of test message. The message was sent at: $d"
        }

        B
        {
            return "XJS*C4JDBQADN1.NSBN3*2IDNEN*GTUBE-STANDARD-ANTI-UBE-TEST-EMAIL*C.34X"
        }

        C
        {
            return "https://www.spamlink.contoso.com"
        }

        D
        {
            return (Read-Host "Enter a Message Body and press enter")
        }

        E
        {
            $MessageBodyPath=[Environment]::GetFolderPath("Desktop")
            [string] $MessageBodyFile = Read-Host "Enter the target EML file full name with extension
for example : EmlFile.eml
Note: Only eml file format can be parsed`r`n
Full File Name"
            [bool]$isPathValid = Test-Office365RelayScriptItemPath([string] "$MessageBodyPath\$MessageBodyFile")
            if($isPathValid)
            {   
                $emlContent = Get-Content "$MessageBodyPath\$MessageBodyFile" -Encoding utf8
                [int] $startOfEmlContent = ($emlContent | Select-String '<!DOCTYPE html>').LineNumber - 1
                [int] $endOfEmlContent = ($emlContent | Select-String '</html>').LineNumber - 1

                if ($null -ne $emlContent)
                {
                return $emlContent[$startOfEmlContent..$endOfEmlContent]
                }

                else 
                {
                Write-Host "`r`nNo Content Found in the pointed file`r`n" -ForegroundColor Red
                Get-MessageBody($d)
                }
            }

            else 
            {
                Write-Host "Content file not found, check path and file name are valid"
                Get-MessageBody($d)
            }
        }

        default
        {Get-MessageBody($d)}
    }  
}

<#
Get-MessageAttachment function provides:
-Eicar test
-Custom attachment from local machine
#>
function Get-MessageAttachment()
{
    [string] $MessageAttachmentType = Read-Host "`r`nSelect a type of Attachment`r`n
    A : Local File
    B : No Attachment
    
    Answer"
    
        switch($MessageAttachmentType.ToUpper())
        {
            A
            {
                Write-Host  -ForegroundColor Cyan -NoNewline "Enter the target htm file full name with extension
for example : C:\attachmentfile.csv`r`nAnswer : "
                [string] $MessageAttachmentFile = Read-Host        
                $isPathValid = Test-Office365RelayScriptItemPath([string]$MessageAttachmentFile)
                if($isPathValid)
                {   
                    return "$MessageAttachmentFile"
                }

                else 
                {
                    Write-Host "Attachment file not found, check path and file name are valid"
                    Get-MessageAttachment
                }
            }

            B
            {
                return "noAttachment"
            }

            default
            {
                Get-MessageAttachment
            }
        }
}

<#
SMTP Client Submission function:
-sends SMTP Client Submission email
-tries to identify known error 
-writes DSN in logs.
#>
function Send-ClientSubmission([PSCredential]$Credentials)
{
    $Office365RelayErrorList.Clear()

    [int]$Port = Read-Host "Input open outbound port (25 or 587)" -ErrorAction SilentlyContinue -ErrorVariable +Office365RelayErrorList

    if ( ($null -eq $Office365RelayErrorList[0]) -and (($Port -eq 25) -or ($Port -eq 587)) )
    {
        $d = Get-Date
        [string] $o365RelayMessageBody = Get-MessageBody($d)
        [string] $o365RelayMessageAttachment = Get-MessageAttachment

        write-host "Sending Message..."
        
        switch -wildcard ($o365RelayMessageAttachment)
        { 
            'noAttachment'
            {         
                Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver smtp.office365.com -subject "SMTP Client Submission Email - $d" `
                                    -body $o365RelayMessageBody -Credential $Credentials -UseSsl -BodyAsHtml -port $Port `
                                    -ErrorAction Continue -WarningAction SilentlyContinue -ErrorVariable +Office365RelayErrorList
            }

            '*:\*'
            {
                Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver smtp.office365.com -subject "SMTP Client Submission Email - $d" `
                                    -body $o365RelayMessageBody -Credential $Credentials -Attachments $o365RelayMessageAttachment `
                                    -UseSsl -BodyAsHtml -port $Port -ErrorAction Continue -WarningAction SilentlyContinue `
                                    -ErrorVariable +Office365RelayErrorList
            }
            
            Default 
            {
            Write-Host "Not Implemented yet, please go back to Main Menu or Exit"
            Get-RetryMenu
            }
        }
        

                          
        
        if ($null -ne $Office365RelayErrorList[0])
        {
            Get-ActionPlan("SmtpClientSubmissionFunctionErrors")                   
        }
    
        else
        {
            Write-Host "`r`nEmail Sent Succesfully"
        
            Get-RetryMenu
        }
    }

    else
    {
        Send-ClientSubmission($Credentials)
    }

    
}

<#
SMTP Relay function:
-sends email via Anonymous SMTP Submission
-tries to identify error issue 
-writes DSN in logs.
#>
function Send-SMTPRelay([string] $O365RelaySmarthost)
{
    $Office365RelayErrorList.Clear()

    [int]$Port = 25

    if ($null -eq $Office365RelayErrorList[0])
    {
        write-host "Sending Message..."
        $d = Get-Date
        Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver $O365RelaySmarthost -subject "SMTP Relay Email - $d" `
                                -body "This is the body of test message. The message was sent at: $d" -UseSsl -BodyAsHtml -Port $Port `
                                -ErrorAction Continue -WarningAction SilentlyContinue -ErrorVariable +Office365RelayErrorList
                          
        if($null -ne $Office365RelayErrorList[0])
        {
            Get-ActionPlan("SmtpRelayFunctionErrors")            
        }
        
        else
        {
            Write-Host "`r`nEmail Sent Succesfully"

            Get-RetryMenu
        }
    }
}

function Write-ScriptLog([string] $ErrorType) 
{
    $d = Get-Date
    $TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    "`r`n$FailedAction at $d $TimeZone generated Error:`r`n" + $Office365RelayErrorList[0].Exception.Message | `
                    Out-File -Append "$global:WSPath\Office365RelayLogs\$ErrorType.txt"
}

function Export-Office365RelayHtmlReport {
[CmdletBinding()]
param(
    [System.Collections.ArrayList][Parameter(Mandatory=$false)]$SmtpRelayFunctionErrors,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)]$SmtpClientSubmissionFunctionErrors,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)]$AuthenticationFunctionErrors,
    [System.Collections.ArrayList][Parameter(Mandatory=$false)]$SmarthostFunctionErrors
)
    $TableType = "List"
    [string]$NoFailure = "We have not encountered any failures."
    #Prepare SmtpRelayFunctionErrors for HTML Report
    [string]$SectionTitle = "SMTP Relay/Direct Send"
    [string]$Description = "This section should display failures encountered for Smtp Relay/Direct Send attempts."
    if($SmtpRelayFunctionErrors){
        $SectionTitleColor = "Red"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $SmtpRelayFunctionErrors -TableType $TableType

        $null = $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)
    }
    else{
        $SectionTitleColor = "Green"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "String" -EffectiveDataString $NoFailure
        $null = $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)
    }

    #Prepare SmtpClientSumbissionErrors for HTML Report
    [string]$SectionTitle = "SMTP Client Submission"
    [string]$Description = "This section should display failures encountered for SMTP Client Submission attempts."
    if($SmtpClientSubmissionFunctionErrors){
        $SectionTitleColor = "Red"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $SmtpClientSubmissionFunctionErrors -TableType $TableType

        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }
    else{
        $SectionTitleColor = "Green"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "String" -EffectiveDataString $NoFailure
        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }

    #Prepare AuthenticationFunctionErrors for HTML Report
    [string]$SectionTitle = "Authentication"
    [string]$Description = "This section should display failures encountered during Authentication attempts."
    if($SmtpClientAuthenticationFunctionErrors){
        $SectionTitleColor = "Red"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $SmtpClientAuthenticationFunctionErrors -TableType $TableType

        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }
    else{
        $SectionTitleColor = "Green"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "String" -EffectiveDataString $NoFailure
        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }
    
    #Prepare SmarthostFunctionErrors for HTML Report
    [string]$SectionTitle = "Smarthost"
    [string]$Description = "This section should display failures encountered for Smarthost resolution attempts."
    if($SmarthostFunctionErrors){
        $SectionTitleColor = "Red"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "CustomObject" -EffectiveDataArrayList $SmarthostFunctionErrors -TableType $TableType

        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }
    else{
        $SectionTitleColor = "Green"
        $HTMLReportEntry = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description `
                -DataType "String" -EffectiveDataString $NoFailure
        $script:Office365RelayHTMLReportArray.Add($HTMLReportEntry)|Out-Null
    }

    #Export HTML Report
    Export-ReportToHTML -FilePath $HTMLFilePath -PageTitle "Office 365 Relay Test Tool" -ReportTitle "Office 365 Relay Test Tool" `
    -TheObjectToConvertToHTML $Office365RelayHTMLReportArray
}

<#
Error + Action Plan Central repository
#>
function Get-ActionPlan()
{
[CmdletBinding()]
param(
    [string][Parameter(Mandatory=$true)]$ErrorType
)
    switch($ErrorType)
    {
        "SmtpRelayFunctionErrors"
        {   
            [string] $FailedAction = "SMTP Relay Email Sent As : $O365SendAs  to RcptTo : "+($Office365RelayRecipients[0].ToString())
            $SmtpRelayFunctionError = New-Object -TypeName PSObject
            $SmtpRelayFunctionError | Add-Member -MemberType NoteProperty -Name "Failed Action" -Value $FailedAction
            $SmtpRelayFunctionError | Add-Member -MemberType NoteProperty -Name "Error Message" `
                                                    -Value $($Office365RelayErrorList[0].Exception.Message + $O365RelaySmarthost)
            Write-ScriptLog($ErrorType)

            switch -Wildcard ($Office365RelayErrorList[0].Exception.Message)
            {
                '*Unable to connect to the remote server*'
                {
                    [string]$ConsoleActionPlan = "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
For instructions on how to use telnet you can check the following articles:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                }
                
                '*connected party did not properly respond after a period of time*'
                {
                    [string]$ConsoleActionPlan = "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
For instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                }

                '*5.7.606 Access denied, banned sending IP*'
                {
                    [string]$ConsoleActionPlan = "Public Outbound IP listed on : https://sender.office.com/ , visit
and follow directions to request removal.
For more information, you can visit : http://go.microsoft.com/fwlink/?LinkID=526655
Once the delist process is completed, allow 24 hours for full delist approval and propagation.
Note: We do not provide IP Statistics to support the reason for the block."
                }

                '*5.7.64 TenantAttribution; Relay Access Denied*'
                {
                    [string]$ConsoleActionPlan = "Relay was not allowed, If the intention was to perform SMTP Relay via
Certificate Scoped Inbound OnPrem Connector, this error is expected.
This script cannot use local certificates for mail submission.
If you are using an IP Scoped Inbound OnPrem Connector:
- the email is not attributed to your tenant as expected.
If you are using an IP Scoped Inbound OnPrem Connector 
and
SMTP Relay previously worked from this machine
-Install the Telnet client software before you can run this command. For more information, see
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)
Open CMD and run:
telnet $O365RelaySmarthost 25
If you connected successfully to an Office 365 server, expect to receive a response line similar to this:
220 BN1BFFO11FD038.mail.protection.outlook.com Microsoft ESMTP MAIL Service ready at Mon, 18 Apr 2016 07:36:51 +0000
If the connection is not successful, then the network firewall or Internet Service Provider (ISP) may block port 25.
If the response does not contain ‘mail.protection.outlook.com Microsoft ESMTP MAIL Service’, check firewall configuration.
Type the following command: EHLO FQDN.yourdomain.com, and then press Enter. You should receive the following response:
250-DB3FFO11FD036.mail.protection.outlook.com Hello [IP address]
-Ensure the Outbound IP mentioned in this reply is configured on the Inbound OnPrem Connector used for Relay.
https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-150/dn910993(v=exchg.150)?redirectedfrom=MSDN#using-the-ip-address-of-your-email-server
For general information on setting up Office 365 Relay Methods, check:
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/fix-issues-with-printers-scanners-and-lob-applications-that-send-email-using-off#emails-are-no-longer-being-sent-to-external-recipients
For further information regarding `"5.7.64 TenantAttribution; Relay Access Denied`", you can access:
https://docs.microsoft.com/en-us/exchange/troubleshoot/connectors/office-365-notice"
                }

                '*5.7.1 Service unavailable, Client host*blocked using Spamhaus*'
                {
                    [string]$ConsoleActionPlan = "Follow instructions provided in the above listed DSN for delisting the IP.
This list is not managed by Microsoft and the Delist should be processed between the owner of the static outbound IP and Spamhaus."
                }

                '*4.4.62 Mail sent to the wrong Office 365 region. ATTR35*'
                {
                    [string]$ConsoleActionPlan = "Relay was not allowed, If the intention was to perform SMTP Relay via
Certificate Scoped Inbound OnPrem Connector, this error is expected.
This script cannot use local certificates for mail submission.
If you are using an IP Scoped Inbound OnPrem Connector:
- the email is not attributed to your tenant as expected.
If you are using an IP Scoped Inbound OnPrem Connector 
and
SMTP Relay previously worked from this machine
-Install the Telnet client software before you can run this command. For more information, see
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)
Open CMD and run:
telnet $O365RelaySmarthost 25
If you connected successfully to an Office 365 server, expect to receive a response line similar to this:
220 BN1BFFO11FD038.mail.protection.outlook.com Microsoft ESMTP MAIL Service ready at Mon, 18 Apr 2016 07:36:51 +0000
If the connection is not successful, then the network firewall or Internet Service Provider (ISP) may block port 25.
If the response does not contain ‘mail.protection.outlook.com Microsoft ESMTP MAIL Service’ check firewall configuration
Type the following command: EHLO FQDN.yourdomain.com, and then press Enter. You should receive the following response:
250-DB3FFO11FD036.mail.protection.outlook.com Hello [IP address]
-Ensure the Outbound IP mentioned in this reply is configured on the Inbound OnPrem Connector used for Relay.
https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-150/dn910993(v=exchg.150)?redirectedfrom=MSDN#using-the-ip-address-of-your-email-server
For general information on setting up Office 365 Relay Methods, check:
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/fix-issues-with-printers-scanners-and-lob-applications-that-send-email-using-off#emails-are-no-longer-being-sent-to-external-recipients
For further information regarding `"5.7.64 TenantAttribution; Relay Access Denied`", you can access:
https://docs.microsoft.com/en-us/exchange/troubleshoot/connectors/office-365-notice"
                }

                '*A socket operation was attempted to an unreachable network*'
                {
                    [string]$ConsoleActionPlan = "Check for local Internet Connectivity issues.
Try to perform the test using a different Internet Connection."
                }

                '*The operation has timed out*'
                {
                    [string]$ConsoleActionPlan = "Check for local Internet Connectivity issues.
Try to perform the test using a different Internet Connection."
                }
            
                default
                {
                    [string]$ConsoleActionPlan = "Email Could not be sent. Error was not recognized, recorded in SMTPRelay errors."
                }
            }
            Write-Warning -Message $ConsoleActionPlan
            $SmtpRelayFunctionError | Add-Member -MemberType NoteProperty -Name "Suggested Action Plan" -Value $ConsoleActionPlan
            $SmtpRelayFunctionErrors.Add($SmtpRelayFunctionError)|Out-Null 
        }

        "SmtpClientSubmissionFunctionErrors"
        {   
            [string] $FailedAction = "SMTP Client Submission Email Sent As : $O365SendAs, authenticated as : " `
                                        +$Credentials.UserName+", sent to "+($Office365RelayRecipients[0].ToString())
            $SmtpClientSubmissionFunctionError = New-Object -TypeName PSObject
            $SmtpClientSubmissionFunctionError | Add-Member -MemberType NoteProperty -Name "Failed Action" -Value $FailedAction
            $SmtpClientSubmissionFunctionError | Add-Member -MemberType NoteProperty -Name "Error Message" `
                                                    -Value $($Office365RelayErrorList[0].Exception.Message + $O365RelaySmarthost)

            Write-ScriptLog($ErrorType)
        
            switch -Wildcard ($Office365RelayErrorList[0].Exception.Message)
            {
                '*Unable to connect to the remote server*'
                {
                    [string]$ConsoleActionPlan = "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
For instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019"
                }

                '*connected party did not properly respond after a period of time*'
                {
                    [string]$ConsoleActionPlan = "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
`r`nFor instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                }

                '*SendAsDeniedException.MapiExceptionSendAsDenied*'
                {
                    [string]$ConsoleActionPlan = "The authenticated account $ScriptLogAutenticatedUser does not have required SendAs permission for $O365SendAs.
Avoid using a single mailbox with Send As permissions for all your users.
This method is not supported because of complexity and potential issues.
If you find yourself in this unsupported scenario,
Direct Send / SMTP Relay should be used as alternatives.
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3#compare-the-options"
                }

                '*5.7.57 SMTP; Client was not authenticated to send anonymous mail*'
                {
                    [string]$ConsoleActionPlan =  "Smtp Client Authentication may be disabled either for the account or on an organization level in EXO.
These settings can be checked via EXO Powershell :
Get-CasMailbox user@contoso.com|fl SmtpClientAuthenticationDisabled
Get-TransportConfig|fl SmtpClientAuthenticationDisabled
Note: To selectively enable authenticated SMTP for specific mailboxes only:
    -disable authenticated SMTP at the organizational level ($true)
    -enable it for the specific mailboxes ($false)
    -and leave the rest of the mailboxes with their default value ($null).

Be aware the Azure Security Defaults will automatically disable Legacy Authentication
and prevent SMTP Submission from being accepted.
For more information on Azure Security defaults, check public article:
https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/concept-fundamentals-security-defaults"
                }

                default
                {
                    [string]$ConsoleActionPlan =  "Email Could not be sent. Error was not recognized, recorded in Client Submission errors."
                }
            }
            Write-Warning -Message $ConsoleActionPlan
            $SmtpClientSubmissionFunctionError | Add-Member -MemberType NoteProperty -Name "Suggested Action Plan" -Value $ConsoleActionPlan
            $SmtpClientSubmissionFunctionErrors.Add($SmtpClientSubmissionFunctionError)|Out-Null
        }

        "AuthenticationFunctionErrors"
        {   
            [string] $FailedAction = "Authentication with account "+$O365SenderCred.UserName+" failed."
            $SmtpClientAuthenticationFunctionError = New-Object -TypeName PSObject
            $SmtpClientAuthenticationFunctionError | Add-Member -MemberType NoteProperty -Name "Failed Action" -Value $FailedAction
            $SmtpClientAuthenticationFunctionError | Add-Member -MemberType NoteProperty -Name "Error Message" `
                                                    -Value $Office365RelayErrorList[0].Exception.Message            
            Write-ScriptLog($ErrorType)

            [string]$ConsoleActionPlan = "Authentication failed, please test authentication via browser
If this account is MFA enabled, please use an App Password instead of the Password to Authenticate.
You can create a new App Password as shown in the following article:
https://support.office.com/en-us/article/Create-an-app-password-for-Office-365-3e7c860f-bda4-4441-a618-b53953ee1183"
            Write-Warning -Message $ConsoleActionPlan
            $SmtpClientAuthenticationFunctionError | Add-Member -MemberType NoteProperty -Name "Suggested Action Plan" -Value $ConsoleActionPlan
            $SmtpClientAuthenticationFunctionErrors.Add($SmtpClientAuthenticationFunctionError)|Out-Null
        }

        "SmarthostFunctionErrors"
        {   
            [string] $FailedAction = "MX record lookup for smarthost "+$InitialDomain
            $SmarthostFunctionError = New-Object -TypeName PSObject
            $SmarthostFunctionError | Add-Member -MemberType NoteProperty -Name "Failed Action" -Value $FailedAction
            $SmarthostFunctionError | Add-Member -MemberType NoteProperty -Name "Error Message" `
                                                    -Value $Office365RelayErrorList[0].Exception.Message
            Write-ScriptLog($ErrorType)

            [string]$ConsoleActionPlan = "Smarthost lookup failed for $InitialDomain
Are the MX and corresponding A records propagated across internet when checking via 3rd party lookup tools online?
    -Yes : Investigate possible local DNS/Network issues.
    -No : Engage Microsoft Support to resolve the DNS record propagation."
            Write-Warning -Message $ConsoleActionPlan
            $SmarthostFunctionError | Add-Member -MemberType NoteProperty -Name "Suggested Action Plan" -Value $ConsoleActionPlan
            $SmarthostFunctionErrors.Add($SmarthostFunctionError)
        }
    }

    Get-RetryMenu
}

<#
Exit-ScriptAndSaveLogs:
-stops transcript
-displays log file location
-exits the script
#>
function Exit-ScriptAndSaveLogs() 
{
    [string] $logFileLocationNotification = "`r`nAll logs have been saved to the following location: $global:WSPath\Office365RelayLogs `r`n"
    Stop-Transcript
    Export-Office365RelayHtmlReport -SmtpRelayFunctionErrors $SmtpRelayFunctionErrors -SmtpClientSubmissionFunctionErrors $SmtpClientSubmissionFunctionErrors `
                                    -AuthenticationFunctionErrors $SmtpClientAuthenticationFunctionErrors -SmarthostFunctionErrors $SmarthostFunctionErrors

    Write-Host $logFileLocationNotification -ForegroundColor Green
    Write-Host "You will return to O365Troubleshooters MainMenu"
    Read-Key
    Clear-Host
    Start-O365TroubleshootersMenu
}

<#
Retry Function will follow successful email submissions.
It will offer quit option with Stop Transcript. 
    -Rerun Get-MainMenu Menu
    -Stop Transcript
    -Display Log Path
    -Exit
#>
function Get-RetryMenu()
{
    $RetryInput = Read-Host "`r`nPress ENTER to reach Main Menu or Q to exit the script"
                    
    switch($RetryInput)
    {
        default {Get-MainMenu}
        Q       {Exit-ScriptAndSaveLogs}
    }
}


<#
Get-CashedOrFreshCredentials provides:
-option to reuse credentials from previous email submission
-option to input fresh credentials
#>
function Get-CashedOrFreshCredentials()
{
	if ($null -eq $Credentials)
    { return [PSCredential]$Credentials = Get-AuthenticationCredentials }
    
	else 
    {
		$ReUseCredentials = Read-Host "Do you want to use the same Credentials?
Y : Yes
N : No
Answer"

		switch ($ReUseCredentials) 
		{
		    Y { [PSCredential]$Credentials }
	        N { return Get-AuthenticationCredentials }
		    default { Get-CashedOrFreshCredentials }
		}
	}
}

<#
MainMenu function:
-reads RelayMethod input from Console
-writes input at Relay Method selection to a file
-calls required functions for each relay method (may be moved to each corresponding relay function instead)
-calls Exit-ScriptAndSaveLogs for quit from script
#>
function Get-MainMenu()
{
    Clear-Host
    
    $RelayMethod = Read-Host "Office 365 Relay Menu`r`n
A : Client Submission (Basic Authentication)
B : SMTP Relay / Direct Send
Q : Quit Script`r`n
Answer"

    "RuntimeRelayMethodInput#$RuntimeRelayMethodCounter $RelayMethod"|Out-File -Append "$global:WSPath\Office365RelayLogs\ChoicesAtRuntime.txt"
    $RuntimeChoiceCounter++

    switch($RelayMethod.ToUpper())
    {
        A 
        {   
            $Credentials = Get-CashedOrFreshCredentials
            [string]$O365SendAs = Get-ValidEmailAddress("From Email Address")
            [string[]]$Office365RelayRecipients = Get-Office365RelayRecipients
            Send-ClientSubmission($Credentials)
        }
    
        B 
        {
            [string]$O365RelaySmarthost = Find-O365RelaySmarthost
            [string]$O365SendAs = Get-ValidEmailAddress("From Email Address")
            [string[]]$Office365RelayRecipients = Get-Office365RelayRecipients
            Send-SMTPRelay($O365RelaySmarthost)
        }
    
        Q
        {Exit-ScriptAndSaveLogs}
    
        default {Get-MainMenu}
    }
}
#endregion Office365RelayDependencies


function Start-Office365Relay {
	[CmdletBinding()]
    param ()

    Clear-Host

    $SendMailMessageDisclaimer ="
Warning : This Script is only recommended for testing purposes as it uses 'Send-MailMessage' cmdlet, 
which is currently considered obsolete.

This cmdlet does not guarantee secure connections to SMTP servers. While there is no immediate 
replacement available in PowerShell, we recommend you do not use Send-MailMessage at this time. 
See https://aka.ms/SendMailMessage for more information.`r`n"

    Write-Host $SendMailMessageDisclaimer -ForegroundColor Yellow

    Read-Key
    
    $ts = Get-Date -Format yyyyMMdd_HHmm

    #Implement check if Log Folder already exists and provide alternative
    #Write-Host "Created Directories on Desktop:"
    if(!(Test-Office365RelayScriptItemPath("$global:WSPath\Office365Relaylogs"))){
        mkdir "$global:WSPath\Office365RelayLogs"|Out-Null
    }
    Start-transcript -Path "$global:WSPath\Office365RelayLogs\RelayTranscript_$ts.txt"|Out-Null

    [string]$script:HTMLFilePath = "$global:WSPath\Office365RelayLogs\Office365RelayHTMLReport.html"

    $script:RuntimeChoiceCounter = 1
    $script:Office365RelayErrorList = @()
    [System.Collections.ArrayList]$script:SmtpRelayFunctionErrors = @()
    [System.Collections.ArrayList]$script:SmtpClientSubmissionFunctionErrors = @()
    [System.Collections.ArrayList]$script:SmtpClientAuthenticationFunctionErrors = @()
    [System.Collections.ArrayList]$script:SmarthostFunctionErrors = @()
    [System.Collections.ArrayList]$script:Office365RelayHTMLReportArray = @()
    Get-MainMenu
}
