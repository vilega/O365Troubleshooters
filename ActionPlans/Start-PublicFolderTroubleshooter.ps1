##Build a menu for PF T.S Diags starting with 
##     1-PF enviroment general stats (PF Overview,PFs count,MEPFs count,PF MBXs count,DBEB status,explict user permissions added on PF MBXs,Hierarchy sync,autosplit status) 
##     2-T.S 554 5.2.2 mailbox full NDR 
##     3-Diagnose PF dumpster for cases related to either delete items inside PF or PF as a whole
##     4-T.S modern\legacy PFs access in a remote enviroment 
##     5-T.S error Mailbox must be accessed as owner. Owner: f4e1a3f9-a0f3-421f-9340-76a8e561c1d1; Accessing user: /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=a39e513cc49849209ea864396c71dc15-exo1
##
##
######################################################################################
<# 1st requirement install the module O365 TS
Import-Module C:\Users\haembab\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>
Clear-Host
#region Connecting to EXO

$Workloads = "exo"
try {
    Connect-O365PS $Workloads 
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
catch {
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Failure"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
#endregion Connecting to EXO

#region Create working folder for Groups Diag
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\PublicFolderTroubleshooter_$ts"
mkdir $ExportPath -Force |out-null
#endregion Create working folder for Groups Diag

##Initialize HTML Object
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()


##T.S 554 5.2.2 mailbox full NDR
Function Start-MEPFNDRDiagnosis{
    Param(
        [parameter(Mandatory=$true)]
        [String]$MEPFSMTP)   

        [string]$Mitigatemessage="Please ensure to follow the fix to mitigate your issue!"
#region Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
if($ContentPFMBXSizeinB -ge $ContentPFMBXProhibitSendReceiveQuotainB)
{
$Orgreached= "Content Public Folder mailbox hosting that mail-enabled public folder has reached its quota! "
#$UserAction=Read-Host "Do you wish to investigate further by checking if Autosplit has processed that mailbox?`nType Y(Yes) to proceed or N(No) to exit!"
<#if ($UserAction -like "*y*")
{
#Call FIX function
Diagnose-MEPFNDRCause("ContentPFMBXfull")
}#>
[string]$SectionTitle = "Validating against content public folder mailbox quota"
[string]$Description = "Checking if the content public folder mailbox hosting the mail-enabled public folder has reached its quota! "

try {
    $fix=Repair-MEPFNDRCause("ContentPFMBXfull") -ErrorAction stop
    $Description=$Description+$Orgreached+"<br>"+$fix
    [PSCustomObject]$PFMBXContentQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Mitigatemessage
    $null = $TheObjectToConvertToHTML.Add($PFMBXContentQuotareachedHTML)
    $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for ContentPFMBXfull reason"
    $CurrentDescription = "Success"
    write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for ContentPFMBXfull reason"
    $CurrentDescription = "Failure"
    write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}

}
else {
    [string]$SectionTitle = "Validating against content public folder mailbox quota"
    [string]$Description = "Checking if the content public folder mailbox hosting the mail-enabled public folder has reached its quota! "
    [PSCustomObject]$PFMBXContentQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($PFMBXContentQuotareachedHTML)
        
}
#endregion Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
#region Validating individual/Organization public folder quota
##Validate if DefaultPublicFolderProhibitPostQuota at the organization level applies
if($MEPFProperties.ProhibitPostQuota -eq "unlimited")
{
[string]$SectionTitle = "Validating against organization public folder post quota"
[string]$Description = "Checking if public folder total size has reached organization public folder DefaultPublicFolderProhibitPostQuota value! "
##catch unlimited value

##Test to use foldersize or stick to the below
##Validate that MEPF size is < 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
    if($MEPFTotalSizeinB -ge $DefaultPublicFolderProhibitPostQuotainB -and $MEPFTotalSizeinB -le 21474836480)
    {
    $Orgreached= "MEPF size ($MEPFTotalSizeinB Bytes) reached the organization DefaultPublicFolderProhibitPostQuota ($DefaultPublicFolderProhibitPostQuotainB Bytes!)"
    ###Call FIX function
    <#$UserAction=Read-Host "Do you wish to mitigate the issue by increasing the DefaultPublicFolderProhibitPostQuota & DefaultPublicFolderIssueWarningQuota values?`nType Y(Yes) to proceed or N(No) to exit!"
        if ($UserAction -like "*y*")
        {
        Debug-MEPFNDRCause("OrgProhibitPostQuotaReached")        
        }
        #>

        try {
            $fix=Repair-MEPFNDRCause("OrgProhibitPostQuotaReached") -ErrorAction stop
            $Description=$Description+$Orgreached+"<br>"+$fix
            #[PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
            [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Mitigatemessage 
            $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
            $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for OrganizationProhibitPostQuotaReached reason"
            $CurrentDescription = "Success"
            write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
        }
        catch {
            $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for OrganizationProhibitPostQuotaReached reason"
            $CurrentDescription = "Failure"
            write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
        }
      
    }
    elseif($MEPFTotalSizeinB -ge $DefaultPublicFolderProhibitPostQuotainB -and $MEPFTotalSizeinB -ge 21474836480)
    {
    ##Validate that MEPF size is > 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
    $Orgreached= "Mail-enabled public folder size is > 20 GB and reached the Organization DefaultPublicFolderProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
    [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
    $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)

    }
    elseif($MEPFTotalSizeinB -le $DefaultPublicFolderProhibitPostQuotainB -and $MEPFTotalSizeinB -le 21474836480)
    {
    ##Validate that MEPF size is > 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
    #$Orgreached= "No Issue found.`nMail-enabled public folder size is < 20 GB  AND LOWER than Organization DefaultPublicFolderProhibitPostQuota!"
    [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
    }
    else
    {
    $Orgreached= "No Issue found.`nMail-enabled public folder size is > 20 GB  AND LOWER than public folder Organzation DefaultPublicFolderProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
    [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
    $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
    }
    [string]$SectionTitle = "Validating against individual public folder post quota"
    [string]$Description = "Checking if public folder total size has reached individual public folder ProhibitPostQuota value!"
    [PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
}
else
{
    [Int]$script:MEPFProhibitPostQuotainB=$MEPFProperties.ProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
    [string]$SectionTitle = "Validating against organization public folder post quota"
    [string]$Description = "Checking if public folder total size has reached organization public folder DefaultPublicFolderProhibitPostQuota value!"
    [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
    [string]$SectionTitle = "Validating against individual public folder post quota"
    [string]$Description = "Checking if public folder total size has reached individual public folder ProhibitPostQuota value! "

##Validate that MEPF size is < 20 GB AND greater than Individual ProhibitPostQuota
if($MEPFTotalSizeinB -ge $MEPFProhibitPostQuotainB -and $MEPFTotalSizeinB -le 21474836480)
{
$Orgreached="The individual public folder post quota (ProhibitPostQuota $MEPFProhibitPostQuotainB Bytes) has been reached!"
<###Call FIXES function
$UserAction=Read-Host "Do you wish to mitigate the issue by increasing the public folder ProhibitPostQuota value?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
    Debug-MEPFNDRCause("IndProhibitPostQuotaReached")
}
#>

try {
    $fix=Repair-MEPFNDRCause("IndProhibitPostQuotaReached") -ErrorAction stop
    $Description=$Description+$Orgreached+"<br>"+$fix
    $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for IndividualProhibitPostQuotaReached reason"
    [PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Mitigatemessage
    $null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
    $CurrentDescription = "Success"
    write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Running repair function across affected mail enabld public folder $MEPFSMTP for IndividualProhibitPostQuotaReached reason"
    $CurrentDescription = "Failure"
    write-log -Function "Repair affected mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}
}
##Validate that MEPF size is > 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
elseif($MEPFTotalSizeinB -ge $MEPFProhibitPostQuotainB -and $MEPFTotalSizeinB -ge 21474836480)
{
$Orgreached= "The individual public folder post quota (ProhibitPostQuota $MEPFProhibitPostQuotainGB GB) has been reached.`nMail-enabled public folder size ($MEPFTotalSizeinGB GB) is > 20 GB, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
[PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
$null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
}
elseif($MEPFTotalSizeinB -le $MEPFProhibitPostQuotainB -and $MEPFTotalSizeinB -le 21474836480)
{
$Orgreached= "No issue found!" 
[PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
$null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
}
else
{
write-host "Mail-enabled public folder size ($MEPFTotalSizeinGB GB) is > 20 GB and didn't reach public folder ProhibitPostQuota ($MEPFProhibitPostQuotainGB GB) value, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
[PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
$null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
}
}
}


#endregion Validating individual/Organization public folder quota

##Present FIX for 554 5.2.2 mailbox full NDR
Function Repair-MEPFNDRCause
{
 Param(
 [parameter(Mandatory=$true)]
 [String]$Cause 
 )
 [string]$SectionTitle = "<br>"+'<font size="+2"><b>FIX</b></font>'+"<br>"
 if($Cause -eq "OrgProhibitPostQuotaReached")
 {
 $article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-organizationconfig" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-organizationconfig</a>'
 [string]$Description = "Please insert a new Organization DefaultPublicFolderProhibitPostQuota value in correlation with a new DefaultPublicFolderIssueWarningQuota value ensuring that these values are greater than MEPF size($MEPFTotalSizeinB Bytes) using command Set-OrganizationConfig."+"<br>"+"For more information please check the following article: $article"
 #[PSCustomObject]$OrgquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $OrgProhibitQuotaReached
 #$null = $TheObjectToConvertToHTML.Add($OrgquotaHTML)
 }
 
 if($Cause -eq "IndProhibitPostQuotaReached")
 {
    if($MEPFProperties.IssueWarningQuota -like "unlimited")
    {
        if ($MEPFTotalSizeinB -le $DefaultPublicFolderProhibitPostQuotainB -and $MEPFTotalSizeinB -le $DefaultPublicFolderIssueWarningQuotainB) 
        {
            [string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder</a>'
            [string]$Description = "Please set public folder ProhibitPostQuota value to Unlimited to inherit from Organization setting(DefaultPublicFolderProhibitPostQuota $DefaultPublicFolderProhibitPostQuotainB Bytes) or set a new public folder ProhibitPostQuota value ensuring that it's greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"
         #   [PSCustomObject]$IndquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $IndProhibitQuotaReached
          #  $null = $TheObjectToConvertToHTML.Add($IndquotaHTML)
        }
        else {
    #Either increase ProhibitPostQuota by value or set it to unlimited considering is lower than organization configuration value
    [string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder</a>'
    [string]$Description = "Please set a new public folder ProhibitPostQuota value ensuring that it's greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"    
    #[PSCustomObject]$IndquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $IndProhibitQuotaReached
    #$null = $TheObjectToConvertToHTML.Add($IndquotaHTML)
}
    }
    else
    {
    if ($MEPFTotalSizeinB -le $DefaultPublicFolderProhibitPostQuotainB -and $MEPFTotalSizeinB -le $DefaultPublicFolderIssueWarningQuotainB) {
    [string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder</a>'
    [string]$Description = "Please set public folder ProhibitPostQuota\IssueWarningQuota values to Unlimited to inherit from Organization setting(DefaultPublicFolderProhibitPostQuota $DefaultPublicFolderProhibitPostQuotainB\DefaultPublicFolderIssueWarningQuota $DefaultPublicFolderIssueWarningQuotainB Bytes) or set a new public folder ProhibitPostQuota\IssueWarningQuota values ensuring that they are greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"
    #[PSCustomObject]$IndquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $IndProhibitQuotaReached
    #$null = $TheObjectToConvertToHTML.Add($IndquotaHTML)
    }
    else {
            [string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-publicfolder</a>'
            [string]$Description = "Please set a new public folder ProhibitPostQuota\IssueWarningQuota values ensuring that they are greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"
            #[PSCustomObject]$IndquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $IndProhibitQuotaReached
            #$null = $TheObjectToConvertToHTML.Add($IndquotaHTML)
    }
}
}
 
 if($Cause -eq "ContentPFMBXfull")
 {
##add condition Check prohibitsendquota if it was set to a lower value (up to 90 GB)
if($ContentPFMBXProhibitSendReceiveQuotainB -le 96636764160 -and $ContentPFMBXProhibitSendReceiveQuotainB -le $MEPFTotalSizeinB)
{   
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailbox" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailbox</a>'
[string]$Description = "Please ensure to use default value of ProhibitSendReceiveQuota(100 GB) or use a higher value than $ContentPFMBXProhibitSendReceiveQuotainB Bytes for content public folder mailbox <b>$MEPFcontentmailbox</b> using set-mailbox command."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$ContentPFMBXProhibitSendReceiveQuotainBHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($ContentPFMBXProhibitSendReceiveQuotainBHTML)    

}
else {
    try {
        $PublicFolderMailboxDiagnostics=Get-PublicFolderMailboxDiagnostics $MailPublicFolder.contentmailbox -ErrorAction stop
        $CurrentProperty = "Retrieving: $MailPublicFolder content mailbox PublicFolderMailboxDiagnostics logs"
        $CurrentDescription = "Success"
        write-log -Function "Retrieve PublicFolderMailboxDiagnostics logs" -Step $CurrentProperty -Description $CurrentDescription
    }
    catch {
        $CurrentProperty = "Retrieving: $MailPublicFolder content mailbox PublicFolderMailboxDiagnostics logs"
        $CurrentDescription = "Failure"
        write-log -Function "Retrieve PublicFolderMailboxDiagnostics logs" -Step $CurrentProperty -Description $CurrentDescription
    }

##Validate Autosplit Halted status
$Autosplitstatus=$PublicFolderMailboxDiagnostics.autosplitinfo.Substring(0,60).split(":")[1].split("")[1]
if($Autosplitstatus -like "Halted")
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "AutoSplit status is Halted for content public folder mailbox <b>$MEPFcontentmailbox</b> so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitHaltedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitHaltedHTML)
}
##Validat if split completed succesfully or failed
elseif($Autosplitstatus -like "SplitCompleted"){
 $MRFValue=($PublicFolderMailboxDiagnostics.autosplitinfo).Split(";")[0].split(" ")[7].split(":")[1]
##validate MRF retry bucket value
if ($MRFValue -like "0")
{
##Validate the date of split
$PublicFolderSplitProcessor=$PublicFolderMailboxDiagnostics.AssistantInfo.ProcessorsState|Where-Object {$_ -like "*PublicFolderSplitProcessor*"}
$DateofPublicFolderSplitProcessor=$PublicFolderSplitProcessor.Split("=")[1]
##Validate Autosplitting process was recent
if((get-date $DateofPublicFolderSplitProcessor) -ge (get-date).AddDays(-7))
{
#Check if DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you might reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.
$DefaultPublicFolderMovedItemRetention=$OrganizationConfig.DefaultPublicFolderMovedItemRetention.Split(":")[0].split(".")[0]
if((get-date $DateofPublicFolderSplitProcessor) -ge (get-date).AddDays(-$DefaultPublicFolderMovedItemRetention))
{
##we might need to lower DefaultPublicFolderMovedItemRetention value to 1 day and invoke mailbox assistant
[string]$Description = @"
Organization DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you still need to reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention 1.00:00:00
Update-PublicFolderMailbox $($MailPublicFolder.contentmailbox)
Check later after couple of hours if the $($MailPublicFolder.contentmailbox) TotalItemSize has reduced by running the below command.`nGet-MailboxStatistics $($MailPublicFolder.contentmailbox)|ft TotalItemSize `nIf the size is reduced, then the issue is fixed and you may set the MovedItemRetention back to old value of $DefaultPublicFolderMovedItemRetention.00:00:00 using below command.`n Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention $DefaultPublicFolderMovedItemRetention.00:00:00
"@
return $Description
#[PSCustomObject]$AutosplitcompletedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitcompletedHTML)
}
else {
    [string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
    [string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for content public folder mailbox <b>$MEPFcontentmailbox</b> for further investigation."+"<br>"+"For more information please refer to the following article: $article"
    #[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
    #$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)        
}
}
##Something other than DefaultPublicFolderMovedItemRetention value prevented items deletion
else
{
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for content public folder mailbox <b>$MEPFcontentmailbox</b> for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
##Autosplit was done more than 7 days ago
else
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for content public folder mailbox <b>$MEPFcontentmailbox</b> for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
else 
{
##Other Autosplit status 
#Autosplit process is in PROGRESS, Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Autosplit process is in PROGRESS for content public folder mailbox <b>$MEPFcontentmailbox</b> so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitHaltedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitHaltedHTML)
}
}
}
[String]$fix=$SectionTitle+"<br>"+$Description+"<br>"
return $fix
}

#region public folder overview
Function Start-PFOverview{
    

    #condition DBEB disabled and PFs are external/local under recommendations part for receing emails externally,configure multiple connectionfilterpolicies and check across default and custom
    
     #region main public folders overview information
     write-host
     Write-Host "Organization Publicfolders Overview`n-----------------------------------"  -ForegroundColor Cyan 
     [string]$SectionTitle = "Introduction"
     [string]$Description = "This report illustrates an overview over the public folder enviroment in EXO, sharing brief useful information about its structure (eg.PF MBXs count) in addition to sharing some health check reports (eg. PF MBX size check) on that!"
     $issuesinhtml='<span style="color: red">issues</span>'
     $Redinhtml='<span style="color: red">red</span>'
     [PSCustomObject]$StartHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate $issuesinhtml reported in $Redinhtml sections in case found!"
     $null = $TheObjectToConvertToHTML.Add($StartHTML)
     try {
        $HostedConnectionFilterPolicy=Get-HostedConnectionFilterPolicy -ErrorAction stop |where-object {$_.IsDefault -eq "True"}
        $CurrentProperty = "Retrieving public folders info"
        $CurrentDescription = "Success"
        write-log -Function "Retrieve public folders info" -Step $CurrentProperty -Description $CurrentDescription    
        $DirectoryBasedEdgeBlockModeStatus=$HostedConnectionFilterPolicy.DirectoryBasedEdgeBlockMode
        [string]$SectionTitle = "Organization Publicfolders Overview"
        [string]$Description = "Collecting all your organization public folders basic information"
        $PFInfo= New-Object PSObject
        if($DirectoryBasedEdgeBlockModeStatus -like "Default")
        {
         Write-Host "DirectoryBasedEdgeBlockModeStatus = Enabled"
         $MEPFAction="Any mail sent to Mail Enabled Public Folders (MEPF) will be dropped at the service network perimeter because DBEB is enabled in the default connection filter policy, 
         so to bypass that please ensure that MEPFs smtp aliases domains are not existing below (the smtp alias DomainType is set to InternalRelay) or file a support case for microsoft to disable DBEB on the whole tenant(Recommended)!"
         $PFInfo|Add-Member -NotePropertyName "Directory Based Edge Block Mode Status" -NotePropertyValue "Enabled"
        }
        else {
            Write-Host "DirectoryBasedEdgeBlockModeStatus = Disabled"
            $PFInfo|Add-Member -NotePropertyName "Directory Based Edge Block Mode Status" -NotePropertyValue "Disabled"
           }
        $Authaccepteddomains=Get-AcceptedDomain -ErrorAction stop |Where-Object{$_.domaintype -eq "Authoritative"}
        $OrganizationConfig=Get-OrganizationConfig -ErrorAction stop
        $PublicFoldersLocation=$OrganizationConfig.PublicFoldersEnabled
        $PublicFolderMailboxes=Get-Mailbox -PublicFolder -ResultSize unlimited -ErrorAction stop
        [Int]$PublicFolderMailboxesCount=($PublicFolderMailboxes).count
        [Int]$MailEnabledPublicFoldersCount=(Get-MailPublicFolder -ResultSize unlimited).count
        $RootPublicFolderMailbox=$OrganizationConfig.RootPublicFolderMailbox.HierarchyMailboxGuid.Guid.ToString()
        Write-Host "PublicFoldersLocation = $PublicFoldersLocation"
        $PFInfo|Add-Member -NotePropertyName "Public Folders Location" -NotePropertyValue $PublicFoldersLocation
        if ($PublicFoldersLocation -eq "Local") {
            $Publicfolders=Get-PublicFolder -Recurse -ResultSize unlimited |Where-Object {$_.Name -notmatch "IPM_SUBTREE"} -ErrorAction stop
            $publicfolderservinghierarchyMBXs=$PublicFolderMailboxes|Where-Object{$_.IsExcludedFromServingHierarchy -like "false" -and $_.IsHierarchyReady -like "true"}
            [Int]$PublicFoldersCount=($Publicfolders).count - 1
            Write-Host "PublicFolderMailboxesCount = $PublicFolderMailboxesCount"
            Write-Host "PublicFolderServingHierarchyMailboxesCount = $($publicfolderservinghierarchyMBXs.name.count)"
            Write-Host "PublicFoldersCount = $PublicFoldersCount"
            Write-Host "RootPublicFolderMailbox = $RootPublicFolderMailbox"
            Write-Host "OrganizationPublicFolderProhibitPostQuota" = $OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[0]
            Write-Host "OrganizationPublicFolderIssueWarningQuota" = $OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[0]
            Write-Host "MailEnabledPublicFoldersCount = $MailEnabledPublicFoldersCount"
            $OrganizationPublicFolderProhibitPostQuota = $OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[0]
            $OrganizationPublicFolderIssueWarningQuota=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[0]
            $PFInfo|Add-Member -NotePropertyName "PublicFolder Mailboxes Count" -NotePropertyValue $PublicFolderMailboxesCount
            $PFInfo|Add-Member -NotePropertyName "PublicFolders Count" -NotePropertyValue $PublicFoldersCount
            $PFInfo|Add-Member -NotePropertyName "Root PublicFolder Mailbox" -NotePropertyValue $RootPublicFolderMailbox
            $PFInfo|Add-Member -NotePropertyName "Organization PublicFolder ProhibitPostQuota" -NotePropertyValue $OrganizationPublicFolderProhibitPostQuota
            $PFInfo|Add-Member -NotePropertyName "Organization PublicFolder IssueWarningQuota" -NotePropertyValue $OrganizationPublicFolderIssueWarningQuota
            $PFInfo|Add-Member -NotePropertyName "MailEnabled PublicFolders Count" -NotePropertyValue $MailEnabledPublicFoldersCount
        }
        else {
            $RemotePublicFolderMailboxes=$OrganizationConfig.RemotePublicFolderMailboxes
            $LockedForMigration=$OrganizationConfig.RootPublicFolderMailbox.LockedForMigration
            if($LockedForMigration -like "True")
            {
                Write-Host "Public folder migration in PROGRESS!" -BackgroundColor Gray 
                Write-Host "PublicFolderMailboxesCount = $PublicFolderMailboxesCount"
                $PFInfo|Add-Member -NotePropertyName "Public folder migration in PROGRESS!"
                $PFInfo|Add-Member -NotePropertyName "PublicFolder Mailboxes Count" -NotePropertyValue $MailEnabledPublicFoldersCount
            }
            else {
                Write-Host "RemotePublicFolderMailboxes = $($RemotePublicFolderMailboxes -join ",")"
                $PFInfo|Add-Member -NotePropertyName "PublicFolder Mailboxes Count" -NotePropertyValue $RemotePublicFolderMailboxes
            }
            
        }

     }
     catch {
        $CurrentProperty = "Retrieving public folders info"
        $CurrentDescription = "Failure"
        write-log -Function "Retrieve public folders info" -Step $CurrentProperty -Description $CurrentDescription
     }
     write-host
     [PSCustomObject]$PFInfoHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject"  -TableType "List" -EffectiveDataArrayList $PFInfo
     $null = $TheObjectToConvertToHTML.Add($PFInfoHTML)
     #endregion main public folders overview information
     #region retrieve publicfolderservinghierarchyMBXs and check if rootPF MBX is serving hierarchy
     #$publicfolderservinghierarchyMBXs=$PublicFolderMailboxes|Where-Object{$_.IsExcludedFromServingHierarchy -like "false" -and $_.IsHierarchyReady -like "true"}
     Write-Host "Public folder mailboxes serving hierarchy: " -NoNewline -ForegroundColor Black -BackgroundColor Yellow
     $publicfolderservinghierarchyMBXs|Format-Table -Wrap -AutoSize  Name,Alias,Guid,ExchangeGuid
     [string]$SectionTitle = "Public folder mailboxes serving hierarchy"
     [string]$Description = "This section illustrates information about public folder mailboxes serving PF hierarchy to end-users" 
     $PFServMBXs=@()
     $PFServMBXs=$publicfolderservinghierarchyMBXs|Select-Object  Name,Alias,Guid,ExchangeGuid
     [PSCustomObject]$PFServMBXsHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $PFServMBXs -TableType "Table"
     $null = $TheObjectToConvertToHTML.Add($PFServMBXsHTML)
     #endregion retrieve publicfolderservinghierarchyMBXs and check if rootPF MBX is serving hierarchy
     #region add check if primary PF MBX doesn't contain content nor serve hierachy to regular MBXs
     write-host "Recommendations:`n================" -ForegroundColor Cyan
     Write-Host
     [string]$SectionTitle = "Primary public folder mailbox diagnosis"
     [string]$Description = "This section illustrates a health check on the root public folder mailbox checking if it's serving PF hierarchy to End-users or used as a content PF MBX" 
     [string]$healthcheck1=""
     [string]$healthcheck2=""
     $healthchecks=New-Object PSObject
     try {
        $PFMBXname=Get-Mailbox -PublicFolder $RootPublicFolderMailbox -ErrorAction stop
     $UserswithrootPFMBXcount=(Get-Mailbox -ResultSize unlimited -ErrorAction stop |Where-Object {$_.EffectivePublicFolderMailbox -Like $PFMBXname.Name}).name.count
     }
    catch {
        ##TODO log the error
    }
     $Checkifrootpfmbxservehierarchy=($publicfolderservinghierarchyMBXs|Where-Object {$_.ExchangeGuid -Like $RootPublicFolderMailbox}).name.count
     if ($Checkifrootpfmbxservehierarchy -ge 1-or $UserswithrootPFMBXcount -ge 1) 
     {
         #$publicfolderservinghierarchyMBXs|Where-Object {$_.ExchangeGuid -Like $RootPublicFolderMailbox}|Format-Table -Wrap -AutoSize Name,Alias,Guid,ExchangeGuid
         $healthcheck1="fail"
     }
     else {
         
         #Write-host "Root public folder mailbox is not used to serve hierachy" -ForegroundColor Green
         $healthcheck1="success"
     }
     [Int]$PFsonrootPFMBXcount=[Int]($Publicfolders|Where-Object {$_.ContentMailboxGuid -Like $RootPublicFolderMailbox}).name.count
     if ($PFsonrootPFMBXcount -eq 0)
     {
         #Write-host "RootPublicFolderMailbox is not hosting content of Public folders" -ForegroundColor Green
         $healthcheck2="success"
     }
     else {
         $healthcheck2="fail"
     }
    if($healthcheck1 -match "fail" -and $healthcheck2 -match "fail")
    {    
        Write-Host "Primary public folder mailbox diagnosis:" -ForegroundColor Black -BackgroundColor Red
        #List the endusers count served by root PF MBX
        Write-Host $UserswithrootPFMBXcount" user(s) found served by primary public folder mailbox, It's not recommended to use root public folder mailbox to serve hierarchy!"
        #List the PFs count hosted on root PF MBX
        Write-host "RootPublicFolderMailbox is hosting content of $PFsonrootPFMBXcount Public folder(s),it's recommended to stop creating public folders hosted on the primary public folder mailbox!"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not used to serve hierachy" -NotePropertyValue "False"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not hosting content of Public folders" -NotePropertyValue "False"
        [PSCustomObject]$RootPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $healthchecks -TableType "List"
        $null = $TheObjectToConvertToHTML.Add($RootPFMBXHTML)

    }
    elseif($healthcheck1 -match "success" -and $healthcheck2 -match "fail"){
        Write-Host "Primary public folder mailbox diagnosis:" -ForegroundColor Black -BackgroundColor Red
        #List the PFs count hosted on root PF MBX
        Write-host "RootPublicFolderMailbox is hosting content of $PFsonrootPFMBXcount Public folder(s),it's recommended to stop creating public folders hosted on the primary public folder mailbox!"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not used to serve hierachy" -NotePropertyValue "True"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not hosting content of Public folders" -NotePropertyValue "False"
        [PSCustomObject]$RootPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $healthchecks -TableType "List"
        $null = $TheObjectToConvertToHTML.Add($RootPFMBXHTML)
    }
    elseif($healthcheck1 -match "fail" -and $healthcheck2 -match "success"){
        Write-Host "Primary public folder mailbox diagnosis:" -ForegroundColor Black -BackgroundColor Red
        #List the endusers count served by root PF MBX
        Write-Host $UserswithrootPFMBXcount" user(s) found served by primary public folder mailbox, It's not recommended to use root public folder mailbox to serve hierarchy!"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not used to serve hierachy" -NotePropertyValue "False"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not hosting content of Public folders" -NotePropertyValue "True"
        [PSCustomObject]$RootPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $healthchecks -TableType "List"
        $null = $TheObjectToConvertToHTML.Add($RootPFMBXHTML)
    }
    else {
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not used to serve hierachy" -NotePropertyValue "True"
        $healthchecks|Add-Member -NotePropertyName "Root PublicFolder Mailbox is not hosting content of Public folders" -NotePropertyValue "True"
        [PSCustomObject]$RootPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $healthchecks -TableType "Table"
        $null = $TheObjectToConvertToHTML.Add($RootPFMBXHTML)
    }
     #endregion add check if primary PF MBX doesn't contain content nor serve hierachy to regular MBXs
     #region add health quota check on PF MBXs 
     [Int]$unhealthyPFMBXcount=0
     $unhealthyPFMBX=@()
     #[Int]$percent=0
     foreach($PublicFolderMailbox in $PublicFolderMailboxes)
     {
        # Write-Progress -Activity "Validating quota on PF MBXs" -Status "$(($percent/$PublicFolderMailboxes.count)*100)% Complete:" -PercentComplete (($percent/$PublicFolderMailboxes.count)*100)
         $PublicFolderMailboxSendReceiveQuota= $PublicFolderMailbox.ProhibitSendReceiveQuota.Split("(")[1].split(" ")[0].Replace(",","")
         try {
             $PublicFolderMailboxMailboxStatistics= Get-MailboxStatistics $PublicFolderMailbox.Guid.Guid.ToString() -ErrorAction stop -WarningAction:SilentlyContinue
             [int]$PFMBXSizeinB=[int]$PublicFolderMailboxMailboxStatistics.TotalItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")+[int]$PublicFolderMailboxMailboxStatistics.TotalDeletedItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")
         }
         catch {
            ##TODOwrite failue in the log
         }
         
         ##Validate PFMBXsize has exceeded 90% PublicFolderMailboxSendReceiveQuota
         if (($PFMBXSizeinB/$PublicFolderMailboxSendReceiveQuota)*100 -ge 90) {
             $unhealthyPFMBXcount++
             $unhealthyPFMBX+=$PublicFolderMailbox
         }
         #$percent++
     }
     
     [string]$SectionTitle = "Public Folder Mailboxes Quota Health Check"
     [string]$Description = "This section illustrates public folder mailboxes that have exceeded autosplit threshold" 
     if($unhealthyPFMBXcount -eq 0)
     {
         write-host
         Write-host "All Public folder mailboxes are on quota healthy state" -ForegroundColor Green
         [string]$PFMBXsQuotaRecommendations="All Public folder mailboxes are on quota healthy state"
         [PSCustomObject]$QuotacheckPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataArrayList $PFMBXsQuotaRecommendations
     }
     else {
         write-host
         Write-host "Please diagnose below public folder mailboxes as their size have exceeded autosplit threshold: " -NoNewline -ForegroundColor Black -BackgroundColor Red
         $unhealthyPFMBX |Format-Table -Wrap -AutoSize Name,Alias,Guid,ExchangeGuid
         $unhealthyPFMBX=$unhealthyPFMBX |Select-Object Name,Alias,Guid,ExchangeGuid |Sort-Object Name
         [PSCustomObject]$QuotacheckPFMBXHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $unhealthyPFMBX -TableType "Table"
     }
     $null = $TheObjectToConvertToHTML.Add($QuotacheckPFMBXHTML)
     #endregion add health quota check on PF MBXs 
     #region MEPF DBEB Check
     [string]$SectionTitle = "Mail-enabled Public Folders Health Check"
     [string]$Description = "This section checks if DBEB will affect MEPFs for receiving external mails" 
     if($null -ne $MEPFAction)
     {  
        Write-Host
        write-host "Mail-enabled Public Folders Health Check:" -ForegroundColor Black -BackgroundColor Red
        Write-Host $MEPFAction
        $Authaccepteddomains|Format-Table Name,DomainName,DomainType
        [PSCustomObject]$MEPFcheckHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "Any mail sent to Mail Enabled Public Folders (MEPF) will be dropped at the service network perimeter because DBEB is enabled in the default connection filter policy, 
        so to bypass that please ensure that MEPFs smtp aliases domains are not existing below (the smtp alias DomainType is set to InternalRelay) or file a support case for microsoft to disable DBEB on the whole tenant(Recommended)!"       
     }
     else {
        [PSCustomObject]$MEPFcheckHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString "DBEB is disabled across the tenant which will allow MEPFs to receive external mails normally"
     }
     $null = $TheObjectToConvertToHTML.Add($MEPFcheckHTML)
     #endregion MEPF Health Check
     #region add health quota check on PFs approaching individual/organization prohibitpostquota,check if we have Giant PFs,that check to be ignored if PFs location is remote
     if ($PublicFoldersLocation -eq "Local") {
         [Int]$unhealthyPFcountapproachingIndQuota=0
         [Int]$unhealthyPFcountapproachingOrgQuota=0
         $unhealthyOrgPF=@()
         $unhealthyIndPF=@()
         $GiantPF=@()
         $UnhealthygiantOrgPF=@()
         $UnhealthygiantIndPF=@()
         foreach($Publicfolder in $Publicfolders)
         {
         [Int64]$OrgPFProhibitPostQuotainB=$OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
        # [Int64]$DefaultPublicFolderIssueWarningQuotainB=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[1].split(" ")[0].Replace(",","")
         $Publicfolderstats=Get-PublicFolderStatistics $($Publicfolder.EntryID)
         [Int64]$PublicfolderTotalSize=[Int64]$Publicfolderstats.TotalItemSize.Split("(")[1].split(" ")[0].Replace(",","")+[Int64]$Publicfolderstats.TotalDeletedItemSize.Split("(")[1].split(" ")[0].Replace(",","")
         ##Check health in regards to Organization quota
             if ($Publicfolder.ProhibitPostQuota -eq "unlimited") {
              if ($PublicfolderTotalSize -ge $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -ge 21474836480) {
                 $unhealthyPFcountapproachingOrgQuota++
                 $UnhealthygiantOrgPF=$UnhealthygiantOrgPF+$publicfolder
                 }
                 elseif ($PublicfolderTotalSize -ge $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -le 21474836480) {
                 $unhealthyPFcountapproachingOrgQuota++
                 $unhealthyOrgPF=$unhealthyOrgPF+$publicfolder

                 }
                 elseif ($PublicfolderTotalSize -le $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -ge 21474836480) {
                     $unhealthyPFcountapproachingOrgQuota++
                     $GiantPF=$GiantPF+$publicfolder
                 }
                 #No action done as it's healthy PF
             }
         ##Check health in regards to PF individual quota
             else {
                 [Int64]$ProhibitPostQuota=$Publicfolder.ProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
                 if ($PublicfolderTotalSize -ge $ProhibitPostQuota -and $PublicfolderTotalSize -ge 21474836480) {
                     $unhealthyPFcountapproachingIndQuota++
                     $UnhealthygiantIndPF=$UnhealthygiantIndPF+$publicfolder
                 }
                  elseif($PublicfolderTotalSize -ge $ProhibitPostQuota -and $PublicfolderTotalSize -le 21474836480) {
                     $unhealthyPFcountapproachingIndQuota++
                     $unhealthyIndPF=$unhealthyIndPF+$publicfolder
                  }
                 elseif ($PublicfolderTotalSize -le $ProhibitPostQuota -and $PublicfolderTotalSize -ge 21474836480) {
                     $unhealthyPFcountapproachingIndQuota++
                     $GiantPF=$GiantPF+$publicfolder
                 }
                 #No action done as it's healthy PF
             }
         }
         if ($unhealthyPFcountapproachingOrgQuota -ge 1)
         {   
             Write-Host
             Write-host "Please diagnose below public folder(s) as their sizes are not compliant for the below reason: " -ForegroundColor Black -BackgroundColor Red
             if($UnhealthygiantOrgPF.Count -ge 1)
             {
                 Write-host "Giant public folder(s) found exceeding OrganizationProhibitPostQuota:`n====================================================================="
                 $UnhealthygiantOrgPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
             }
             if($unhealthyOrgPF.Count -ge 1)
             {
                 Write-host "Public folder(s) found exceeding OrganizationProhibitPostQuota:`n==============================================================="
                 $unhealthyOrgPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
             }
             if($GiantPF.Count -ge 1)
             {
                 Write-host "Giant Public folder(s) found:`n============================="
                 $GiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
             }
         }
         if ($unhealthyPFcountapproachingIndQuota -ge 1)
         {
             Write-Host
             Write-host "Please diagnose below public folder(s) as their sizes are not compliant for the below reason: "  -ForegroundColor Black -BackgroundColor Red
             if($UnhealthygiantIndPF.Count -ge 1)
             {
                 Write-host "Giant public folder(s) found exceeding individual ProhibitPostQuota:`n==================================================================================="
                 $UnhealthygiantIndPF|Format-Table -Wrap -AutoSize Name,Identity,ProhibitPostQuota,EntryID
             }
             if($unhealthyIndPF.Count -ge 1)
             {
                 Write-host "Public folder(s) found exceeding their individual ProhibitPostQuota:`n===================================================================="
                 $unhealthyIndPF|Format-Table -Wrap -AutoSize Name,Identity,ProhibitPostQuota,EntryID
             }
             if($GiantPF.Count -ge 1)
             {
                 Write-host "Giant Public folder(s) found:`n============================="
                 $GiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
             }
         }
         else 
         {
             Write-Host
             Write-Host "All Public folder(s) are on quota healthy state" -ForegroundColor Green
         }
 
     }
     
     #endregion add health quota check on PFs approaching individual/organization prohibitpostquota,check if we have Giant PFs
    <#
     #region ResultReport
        [string]$FilePath = $ExportPath + "\PublicFolderOverview.html"
        Export-ReportToHTML -FilePath $FilePath -PageTitle "Public Folders Overview" -ReportTitle "Public Folders Overview" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
    #Question to ask enduser for opening the HTMl report
    $OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
    if ($OpenHTMLfile -like "*y*")
    {
        Write-Host "Opening report...." -ForegroundColor Cyan
        Start-Process $FilePath
    }
    #endregion ResultReport
    #>
   
# End of the Diag
#Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
#Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu
     
     #TODOcondition for validation on root PFs have dumpsterentryIDs
     #TODO specify HRR MBXs in case exist
     #TODO Think about adding Autosplit status will take so much time for huge enviroments
     #TODO add MEPFs are synced using AD connect
     #TODO add print for report to html
  }
  
 Function ValidatePermission
 {
 Param(
 [parameter(Mandatory=$true)]
 [PSCustomObject]$Perms 
         )  
 [array]$workingpermissions=@("editor","owner","publishingeditor","deleteallitems")
 if ($null -ne $Perms) {
 foreach($perm in $Perms.AccessRights)
     {
         if($workingpermissions.Contains($($perm.ToLower())))
         {
         ##user has the permission skip by break the forloop
         return "user has permission"
         }
     }
     return "user has no permission"
 }
 else {
     return "user has no permission"
 }
 }
 
  Function ValidatePFDumpster{
     Param(
         [parameter(Mandatory=$true)]
         [String]$Pfolder 
         )
 
 #region public folder diagnosis        
 try {
     $Publicfolder=Get-PublicFolder $Pfolder -ErrorAction stop
     $Publicfolderdumpster=Get-PublicFolder $Publicfolder.DumpsterEntryId -ErrorAction stop
     #$Publicfolderdumpster="0000000096CE4B52BB898C4FA11E7E230A3C8EE7010077B56B4D3B88794B9817E41A07D18FF500000000001F0000"
     $pfmbx=Get-mailbox -PublicFolder $Publicfolder.ContentMailboxGuid.Guid
     $PfMBXstats=Get-mailboxStatistics $Publicfolder.ContentMailboxGuid.Guid -ErrorAction stop
     $IPM_SUBTREE=Get-PublicFolder \ -ErrorAction stop
     $NON_IPM_SUBTREE=Get-PublicFolder \NON_IPM_SUBTREE -ErrorAction stop
     $DUMPSTER_ROOT=Get-PublicFolder \NON_IPM_SUBTREE\DUMPSTER_ROOT -ErrorAction stop
     $CurrentProperty = "Retrieving: $($Publicfolder.identity) & its dumpster for diagnosing"
     $CurrentDescription = "Success"
     write-log -Function "Retrieve public folder & its dumpster statistics" -Step $CurrentProperty -Description $CurrentDescription
     [string]$SectionTitle = "Introduction"
     [string]$Description = "This report illustrates causes behind users with sufficient permissions cannot delete items under public folder using OWA\Outlook or cannot remove the entire public folder as a whole." + "<br>"+ "Checks run on Public folder: <b>$($Publicfolder.identity)</b>"
     $blockersinhtml='<span style="color: red">BLOCKERS</span>'
     [PSCustomObject]$StartHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate $blockersinhtml in case found!"
     $null = $TheObjectToConvertToHTML.Add($StartHTML)
 }
 catch {
     $Errorencountered=$Global:error[0].Exception
     $CurrentProperty = "Retrieving: $($Pfolder) & its dumpster for diagnosing"
     $CurrentDescription = "Failure with error: "+$Errorencountered
     write-log -Function "Retrieve public folder & its dumpster statistics" -Step $CurrentProperty -Description $CurrentDescription
     Write-Host "Error encountered during executing the script!"-ForegroundColor Red
     Write-Host $Errorencountered -ForegroundColor Red
     Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
     Start-Sleep -Seconds 3
     Read-Key
     # Go back to the main menu
     Start-O365TroubleshootersMenu
     ##write log and exit function
 }
 #endregion public folder diagnosis
 
 #region to validate permissions across the public folder
 #Identify if item or folder
 $ItemORFolder=Read-Host "Please specify if the issue is related to a user who is not able to delete an item inside a public folder or neither a user with owner permissions nor the admin are not able to delete the Public folder as awhole, Type (I) for Item or (F) for Folder"
 if ($ItemORFolder.ToLower() -eq "i") {
     #validate explict permission & default permission if item
 $Affecteduser=Get-ValidEmailAddress("Please provide an affected user smtp!")
 try {
     $User=Get-Mailbox $Affecteduser -ErrorAction stop
     $Explicitperms=Get-PublicFolderClientPermission $Publicfolder.EntryId -User $User.Guid.Guid.tostring() -ErrorAction SilentlyContinue
     $Defaultperms=Get-PublicFolderClientPermission $Publicfolder.EntryId -User Default -ErrorAction SilentlyContinue
     if($null -ne $Explicitperms)
     {
         $Explicitpermsresult=ValidatePermission($Explicitperms)
     }
     else {
         $Explicitpermsresult="user has no permission"
     }
     if ($null -ne $Defaultperms){
         $Defaultpermsresult=ValidatePermission($Defaultperms)
     }
     else {
         $Defaultpermsresult -match "user has no permission"
     }
     if ($Explicitpermsresult -match "user has no permission" -and $Defaultpermsresult -match "user has no permission") 
     {
         #user has no permission to delete break the script
         [string]$SectionTitle = "Validating User Permission"
         [string]$Description = "Checking if user $($User.PrimarySmtpAddress) has sufficient permissions to delete"   
         [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring "Neither $($User.PrimarySmtpAddress) nor Default user have sufficient permissions to delete items inside $($publicfolder.identity)"
         $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
         #Identify the type of permission that has the trouble ,add the user to the report
     }
     else {
         #user has permission to delete continue with the script
         [string]$SectionTitle = "Validating User Permission"
         [string]$Description = "Checking if user $($User.PrimarySmtpAddress) has sufficient permissions to delete"   
         [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
         $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
     }
     
     
 }
 catch {
     #log the error and quit
     $Errorencountered=$Global:error[0].Exception
     $CurrentProperty = "Validating if user has sufficient permissions to delete"
     $CurrentDescription = "Failure with error: "+$Errorencountered
     write-log -Function "Validate user permissions" -Step $CurrentProperty -Description $CurrentDescription
     Write-Host "Error encountered during executing the script!"-ForegroundColor Red
     Write-Host $Errorencountered -ForegroundColor Red
     Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
     Start-Sleep -Seconds 3
     Read-Key
     # Go back to the main menu
     Start-O365TroubleshootersMenu
     ##write log and exit function
 }
 }
 elseif ($ItemORFolder.ToLower() -eq "f") {
 
     #continue with the script    
 }
 else {
     Write-Host "You didn't provide an expected input!" -ForegroundColor Red
     Write-Host "Relaunching the main menu again" -ForegroundColor Yellow 
     Start-Sleep -Seconds 3
     Read-Key
     # Go back to the main menu
     Start-O365TroubleshootersMenu
 }
 
 #endregion to validate permissions across the public folder
 
 #region to validate content PF MBX across both PF & its dumpster
 if($Publicfolder.ContentMailboxGuid.Guid -ne $Publicfolderdumpster.ContentMailboxGuid.Guid)
 {   
     #raise a support request for microsoft including get-publicfolder logs 
     [string]$SectionTitle = "Validating content public folder mailbox"
     [string]$Description = "Checking if public folder & its dumpster has the same content public folder mailbox"   
     [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring "Please raise a support request for microsoft including the HTML report & compressed logs folder"
     $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
     #add the report to the request + logs folder
 
 }
 else{
     [string]$SectionTitle = "Validating content public folder mailbox"
     [string]$Description = "Checking if public folder & its dumpster has the same content public folder mailbox"   
     [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
     $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 }
 #endregion to validate content PF MBX across both PF & its dumpster
 
 #region to validate EntryId &DumpsterEntryID values are mapped properly 
 if($Publicfolder.EntryId -ne $Publicfolderdumpster.DumpsterEntryID -or $Publicfolder.DumpsterEntryID -ne $Publicfolderdumpster.EntryId)
 {
 #raise a support request for microsoft including get-publicfolder logs 
 [string]$SectionTitle = "Validating public folder EntryId mapping"
 [string]$Description = "Checking if public folder EntryId & DumpsterEntryID values are mapped properly"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring "Please raise a support request for microsoft including the HTML report & compressed logs folder"
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 #add the report to the request +logs folder
 }
 else {
 [string]$SectionTitle = "Validating public folder EntryId mapping"
 [string]$Description = "Checking if public folder EntryId & DumpsterEntryID values are mapped properly"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 }
 #endregion to validate EntryId &DumspterEnryID values are mapped properly
 
 #region to validate public folder mailbox TotalDeletedItemSize value hasn’t reached its RecoverableItemsQuota value
 [Int64]$pfmbxRecoverableItemsQuotainB=[Int64]$pfmbx.RecoverableItemsQuota.Split("(")[1].split(" ")[0].Replace(",","")
 [Int64]$PfMBXstatsinB=[Int64]$PfMBXstats.TotalDeletedItemSize.Value.tostring().Split("(")[1].split(" ")[0].Replace(",","")        
 if($PfMBXstatsinB -ge $pfmbxRecoverableItemsQuotainB  )
 {
 <#
 To resolve a scenario where content public folder mailbox TotalDeletedItemSize value has reached RecoverableItemsQuota value, users could manually clean up the dumpster using:
     Outlook RecoverDeletedItems
     MFCMAPI please refer to following article to check steps related to get to public folder dumpster using MFCMAPI then select unrequired items to be purged permanently
 #>
 $article='<a href="https://docs.microsoft.com/en-us/archive/blogs/exovoice/public-folders-data-recovery-scenarios" target="_blank">article</a>'
 $RecoverDeletedItems='<a href="https://docs.microsoft.com/en-us/exchange/troubleshoot/public-folders/cannot-delete-items-public-folder" target="_blank">RecoverDeletedItems</a>'
 $FixTotalDeletedItemSize="To resolve a scenario where content public folder mailbox TotalDeletedItemSize value has reached RecoverableItemsQuota value, users could manually clean up the dumpster using:<br>
 ->Outlook $RecoverDeletedItems<br>
 ->MFCMAPI please refer to the following $article to check steps related to get to public folder dumpster using MFCMAPI then select unrequired items to be purged permanently"
 [string]$SectionTitle = "Validating TotalDeletedItemSize for content public folder mailbox"
 [string]$Description = "Checking if public folder mailbox TotalDeletedItemSize value has not reached its RecoverableItemsQuota value"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $FixTotalDeletedItemSize
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 }
 else {
 [string]$SectionTitle = "Validating TotalDeletedItemSize for content public folder mailbox"
 [string]$Description = "Checking if public folder mailbox TotalDeletedItemSize value has not reached its RecoverableItemsQuota value"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 }
 #endregion to validate public folder mailbox TotalDeletedItemSize value hasn’t reached its RecoverableItemsQuota value
 
 #region to validate that root public folders “IPM_SUBTREE & NON_IPM_SUBTREE & DUMPSTER_ROOT” DumpsterEntryID values are populated 
 if($null -eq $IPM_SUBTREE.DumpsterEntryId -or $null -eq $NON_IPM_SUBTREE.DumpsterEntryId -or $null -eq $DUMPSTER_ROOT.DumpsterEntryId)
 {
 #raise a support request for microsoft including get-publicfolder for root folder logs 
 [string]$SectionTitle = "Validating root public folders"
 [string]$Description = "Checking if root public folders have DumpsterEntryId value"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring "Please raise a support request for microsoft including the HTML report & compressed logs folder"
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 #add the report to the request +logs folder
 }
 else {
 [string]$SectionTitle = "Validating root public folders"
 [string]$Description = "Checking if root public folders have DumpsterEntryId value"   
 [PSCustomObject]$ConditioncheckPFPermissionhtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
 $null = $TheObjectToConvertToHTML.Add($ConditioncheckPFPermissionhtml)
 }
 #endregion to validate that root public folders “IPM_SUBTREE & NON_IPM_SUBTREE & DUMPSTER_ROOT” DumpsterEntryID values are populated  
 #region ResultReport
 [string]$FilePath = $ExportPath + "\PublicFolderdumpsterTroubleshooter.html"
 Export-ReportToHTML -FilePath $FilePath -PageTitle "Public Folder Troubleshooter" -ReportTitle "Public Folder Dumpster Troubleshooter" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
 #Question to ask enduser for opening the HTMl report
 $OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
 if ($OpenHTMLfile -like "*y*")
 {
 Write-Host "Opening report...." -ForegroundColor Cyan
 Start-Process $FilePath
 }
 #endregion ResultReport
 #create zip file for logs folder
 $tstamp= get-date -Format yyyyMMdd_HHmmss
 if (!(Test-Path  "$ExportPath\logs_$tstamp"))
 {
     mkdir "$ExportPath\logs_$tstamp" -Force |out-null
 }
 $Publicfolder|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\Publicfolder.txt" -NoClobber 
 $Publicfolderdumpster|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\Publicfolderdumpster.txt" -NoClobber
 $pfmbx|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\pfmbx.txt" -NoClobber
 $PfMBXstats|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\PfMBXstats.txt" -NoClobber 
 $IPM_SUBTREE|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\IPM_SUBTREE.txt" -NoClobber
 $NON_IPM_SUBTREE|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\NON_IPM_SUBTREE.txt" -NoClobber
 $DUMPSTER_ROOT|Format-List|Out-File -FilePath "$ExportPath\logs_$tstamp\DUMPSTER_ROOT.txt" -NoClobber
 Compress-Archive -Path "$ExportPath\logs_$tstamp" -DestinationPath $ExportPath\logs_$tstamp

 #relanuching main menu again
 Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
     Start-Sleep -Seconds 3
     Read-Key
     # Go back to the main menu
     Start-O365TroubleshootersMenu
     #write log and exit function
  }
 


##Code for Menu
$PFMenu=@"
1 - Public folder overview
2 - Diagnosing 554 5.2.2 mailbox full NDR received on sending to MEPF
3 - Diagnosing the cause behind the failure of deleting a public folder item or the whole public folder
Q  Quit
     
Select a task by number or Q to quit
"@

$menuchoice=Read-Host $PFMenu
$menuchoice = $menuchoice.ToLower()
if ($menuchoice -eq 1)
{
    #HTML report in next release
    Write-Warning "This diagnostic is going to be generating HTML report output over O365Troubleshooter upcoming release!"
    Write-Warning "This diagnostic is tested on small public folder enviroments (1k) so please expect some delay if you have medium to huge public folder enviroments!" 
    ##TODO allow some time warning for huge enviroments of PFs
    Read-Key
    #Clear-Host
    Start-PFOverview
}
elseif ($menuchoice -eq 2)
{
#region Get the affected MEPF SMTP
$MEPFSMTP=Get-ValidEmailAddress("Email address of the mail enabled public folder ")
#endregion Get the affected MEPF SMTP
#region Intro with group name 
[string]$SectionTitle = "Introduction"
[String]$article='<a href="https://docs.microsoft.com/en-us/exchange/troubleshoot/email-delivery/cannot-send-mail-mepf" target="_blank">Error when sending email to mail-enabled public folders in Office 365: 554 5.2.2 mailbox full</a>'
[string]$Description = "This report illustrates causes behind a non-delivery report (NDR) with error code 554 5.2.2 when sending emails to a mail-enabled public folder: "+"<b>$MEPFSMTP</b>"+", Sections in RED are for checks BLOCKERS while Sections in GREEN are for checks PASSED!"
$Description=$Description+"<br>"+"For more informtion please check: $article"
[PSCustomObject]$StartHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate causes in case found by checking FIX section!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
#endregion Intro with group name    
#region script variables used in functions
try {
    [PSCustomObject]$MailPublicFolder=Get-MailPublicFolder $MEPFSMTP -ErrorAction stop
    [PSCustomObject]$ContentPFMBXStatistics=Get-MailboxStatistics $MailPublicFolder.contentmailbox -ErrorAction stop
    [PSCustomObject]$ContentPFMBXProperties=Get-Mailbox -PublicFolder $MailPublicFolder.contentmailbox -ErrorAction stop
    [PSCustomObject]$OrganizationConfig=Get-OrganizationConfig -ErrorAction stop
    [PSCustomObject]$MEPFStatistics=Get-PublicFolderStatistics -identity $MailPublicFolder.EntryID -ErrorAction stop
    [PSCustomObject]$MEPFProperties=Get-PublicFolder $MailPublicFolder.EntryID -ErrorAction stop
    [string]$MEPFcontentmailbox=$MailPublicFolder.contentmailbox
    [Int64]$MEPFProhibitPostQuotainGB = $MEPFProhibitPostQuotainB /(1024*1024*1024)
    [Int64]$ContentPFMBXProhibitSendReceiveQuotainB=$ContentPFMBXProperties.ProhibitSendReceiveQuota.Split("(")[1].split(" ")[0].Replace(",","")
    [Int64]$ContentPFMBXSizeinB=[Int64]$ContentPFMBXStatistics.TotalItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")+[Int64]$ContentPFMBXStatistics.TotalDeletedItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")
    [Int64]$DefaultPublicFolderProhibitPostQuotainB=$OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
    [Int64]$DefaultPublicFolderIssueWarningQuotainB=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[1].split(" ")[0].Replace(",","")
    [Int64]$script:MEPFTotalSizeinB=[Int64]$MEPFStatistics.TotalItemSize.Split("(")[1].split(" ")[0].Replace(",","")+[Int64]$MEPFStatistics.TotalDeletedItemSize.Split("(")[1].split(" ")[0].Replace(",","")
    [Int64]$MEPFTotalSizeinGB = $MEPFTotalSizeinB /(1024*1024*1024)
    $CurrentProperty = "Retrieving: $MEPFSMTP object properties & statistics, content mailbox properties & statistics AND organization configuration"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve object properties & statistics" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $Errorencountered=$Global:error[0].Exception
    $CurrentProperty = "Retrieving: $MEPFSMTP object properties & statistics, content mailbox properties & statistics AND organization configuration"
    $CurrentDescription = "Failure with error: "+$Errorencountered
    write-log -Function "Retrieve object properties & statistics" -Step $CurrentProperty -Description $CurrentDescription
    Write-Host "Error encountered during executing the script!"-ForegroundColor Red
    Write-Host $Errorencountered -ForegroundColor Red
    Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
    Start-Sleep -Seconds 3
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
}
#endregion global variables used in functions
try {
    Start-MEPFNDRDiagnosis($MEPFSMTP) -ErrorAction Stop
    $CurrentProperty = "Diagnosing: $MEPFSMTP for 554 5.2.2 mailbox full NDR cause"
    $CurrentDescription = "Success"
    write-log -Function "Diagnose mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Diagnosing: $MEPFSMTP for 554 5.2.2 mailbox full NDR cause"
    $CurrentDescription = "Failure"
    write-log -Function "Diagnose mail enabled public folder" -Step $CurrentProperty -Description $CurrentDescription
}

#region ResultReport
[string]$FilePath = $ExportPath + "\PublicFolderTroubleshooter.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "PublicFolderTroubleshooter" -ReportTitle "PublicFolderTroubleshooter" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
#Question to ask enduser for opening the HTMl report
$OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile -like "*y*")
{
Write-Host "Opening report...." -ForegroundColor Cyan
Start-Process $FilePath
}
#endregion ResultReport
}
elseif ($menuchoice -eq 3)
{
    $Pfolder=Read-Host "Please enter the affected public folder identity or EntryID ex.\PF1"
    ValidatePFDumpster($Pfolder)
}

elseif($menuchoice -eq "q")
{
    Write-Host "Quitting...."
    Write-Host "Relaunching the main menu again" -ForegroundColor Yellow 
    Start-Sleep -Seconds 3
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
}
else {
    Write-Host "You didn't provide an expected input!"
    Write-Host "Relaunching the main menu again" -ForegroundColor Yellow 
    Start-Sleep -Seconds 3
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
}
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu




