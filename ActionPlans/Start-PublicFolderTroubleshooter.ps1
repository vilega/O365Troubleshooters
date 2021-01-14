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
Import-Module C:\Users\a-haemb\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
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
            [string]$Description = "Please set public folder ProhibitPostQuota value to Unlimited to inherit from Organization setting or set a new public folder ProhibitPostQuota value ensuring that it's greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"
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
    [string]$Description = "Please set public folder ProhibitPostQuota\IssueWarningQuota values to Unlimited to inherit from Organization setting or set a new public folder ProhibitPostQuota\IssueWarningQuota values ensuring that they are greater than the public folder size($MEPFTotalSizeinB Bytes)using command Set-PublicFolder."+"<br>"+"For more information please check the following article: $article"
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
[string]$Description = "Please ensure to use default value of ProhibitSendReceiveQuota(100 GB) or use a higher value than $ContentPFMBXProhibitSendReceiveQuotainB Bytes using set-mailbox command."+"<br>"+"For more information please refer to the following article: $article"
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
[string]$Description = "AutoSplit status is Halted so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
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
    [string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
    #[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
    #$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)        
}
}
##Something other than DefaultPublicFolderMovedItemRetention value prevented items deletion
else
{
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
##Autosplit was done more than 7 days ago
else
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
else 
{
##Other Autosplit status 
#Autosplit process is in PROGRESS, Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$article='<a href="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps" target="_blank">https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps</a>'
[string]$Description = "Autosplit process is in PROGRESS so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation."+"<br>"+"For more information please refer to the following article: $article"
#[PSCustomObject]$AutosplitHaltedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
#$null = $TheObjectToConvertToHTML.Add($AutosplitHaltedHTML)
}
}
}
[String]$fix=$SectionTitle+"<br>"+$Description+"<br>"
return $fix
}




##Code for Menu
$PFMenu=@"
1 - Public folder overview
2 - Diagnosing 554 5.2.2 mailbox full NDR received on sending to MEPF
Q  Quit
     
Select a task by number or Q to quit
"@

$menuchoice=Read-Host $PFMenu
if ($menuchoice -eq 1)
{
    Start-PFDataCollection
}
if ($menuchoice -eq 2)
{
#region Get the affected MEPF SMTP
$MEPFSMTP=Get-ValidEmailAddress("Email address of the mail enabled public folder ")
#endregion Get the affected MEPF SMTP
#region Intro with group name 
[string]$SectionTitle = "Introduction"
[String]$article='<a href="https://docs.microsoft.com/en-us/exchange/troubleshoot/email-delivery/cannot-send-mail-mepf" target="_blank">Error when sending email to mail-enabled public folders in Office 365: 554 5.2.2 mailbox full</a>'
[string]$Description = "This report illustrates causes behind a non-delivery report (NDR) with error code 554 5.2.2 when sending emails to a mail-enabled public folder: "+$MEPFSMTP+", Sections in RED are for checks BLOCKERS while Sections in GREEN are for checks PASSED!"
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
    $CurrentProperty = "Retrieving: $MEPFSMTP object properties & statistics, content mailbox properties & statistics AND organization configuration"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve object properties & statistics" -Step $CurrentProperty -Description $CurrentDescription
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
else {
    Exit
}




#region public folder overview
Function Start-PFDataCollection{
    
   
    #region main public folders overview information
    write-host
    write-host
    Write-Host "Public Folders Overview`n========================"  -ForegroundColor Cyan
    $HostedConnectionFilterPolicy=Get-HostedConnectionFilterPolicy 
    $DirectoryBasedEdgeBlockModeStatus=$HostedConnectionFilterPolicy.DirectoryBasedEdgeBlockMode
    if($DirectoryBasedEdgeBlockModeStatus -like "Default")
    {
        Write-Host "DirectoryBasedEdgeBlockModeStatus = Enabled"
    }
    else {
        Write-Host "DirectoryBasedEdgeBlockModeStatus = Disabled"
    }
    $OrganizationConfig=Get-OrganizationConfig
    $PublicFoldersLocation=$OrganizationConfig.PublicFoldersEnabled
    $PublicFolderMailboxes=Get-Mailbox -PublicFolder -ResultSize unlimited
    [Int]$PublicFolderMailboxesCount=($PublicFolderMailboxes).count
    [Int]$MailEnabledPublicFoldersCount=(Get-MailPublicFolder -ResultSize unlimited).count
    $RootPublicFolderMailbox=$OrganizationConfig.RootPublicFolderMailbox.HierarchyMailboxGuid.Guid.ToString()
    Write-Host "PublicFoldersLocation = $PublicFoldersLocation"
    if ($PublicFoldersLocation -eq "Local") {
        $Publicfolders=Get-PublicFolder -Recurse -ResultSize unlimited
        [Int]$PublicFoldersCount=($Publicfolders).count - 1
        Write-Host "PublicFolderMailboxesCount = $PublicFolderMailboxesCount"
        Write-Host "PublicFoldersCount = $PublicFoldersCount"
        Write-Host "RootPublicFolderMailbox = $RootPublicFolderMailbox"       
    }
    else {
        $RemotePublicFolderMailboxes=$OrganizationConfig.RemotePublicFolderMailboxes
        $LockedForMigration=$OrganizationConfig.RootPublicFolderMailbox.LockedForMigration
        if($LockedForMigration -like "True")
        {
            Write-Host "Public folder migration in PROGRESS!" -BackgroundColor Gray 
            Write-Host "PublicFolderMailboxesCount = $PublicFolderMailboxesCount"
        }
        else {
            Write-Host "RemotePublicFolderMailboxes = $($RemotePublicFolderMailboxes -join ",")"
        }
        
    }
    Write-Host "MailEnabledPublicFoldersCount = $MailEnabledPublicFoldersCount" 
    #endregion main public folders overview information
    #region retrieve publicfolderservinghierarchyMBXs and check if rootPF MBX is serving hierarchy
    $publicfolderservinghierarchyMBXs=$PublicFolderMailboxes|Where-Object{$_.IsExcludedFromServingHierarchy -like "false" -and $_.IsHierarchyReady -like "true"}
    Write-Host "Public folder hierarchy serving mailboxes: " -NoNewline -ForegroundColor Black -BackgroundColor Yellow
    $publicfolderservinghierarchyMBXs|Format-Table -Wrap -AutoSize Name,Alias,Guid,ExchangeGuid
    #endregion retrieve publicfolderservinghierarchyMBXs and check if rootPF MBX is serving hierarchy
    #region add check if primary PF MBX doesn't contain content nor serve hierachy to regular MBXs
    Write-Host "Root public folder mailbox diagnosis:" -ForegroundColor Black -BackgroundColor Yellow
    if ($publicfolderservinghierarchyMBXs|Where-Object {$_.ExchangeGuid -Like $RootPublicFolderMailbox}) 
    {
        Write-host "It's not recommended to use root public folder mailbox to serve hierarchy!" -ForegroundColor Red -NoNewline
        $publicfolderservinghierarchyMBXs|Where-Object {$_.ExchangeGuid -Like $RootPublicFolderMailbox}|Format-Table -Wrap -AutoSize Name,Alias,Guid,ExchangeGuid
    }
    else {
        Write-host "Root public folder mailbox is not used to serve hierachy" -ForegroundColor Green
    }
    if ([Int]($Publicfolders|Where-Object {$_.ContentMailboxGuid -Like $RootPublicFolderMailbox}).name.count -eq 1)
    {
        Write-host "RootPublicFolderMailbox is not hosting content of Public folders" -ForegroundColor Green
    }
    else {
        Write-host "RootPublicFolderMailbox is hosting content of Public folders,it's recommended to stop creating public folders hosted on the primary public folder mailbox!" -ForegroundColor Red
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
            $PublicFolderMailboxMailboxStatistics= Get-MailboxStatistics $PublicFolderMailbox.Alias -ErrorAction stop -WarningAction:SilentlyContinue
            [int]$PFMBXSizeinB=[int]$PublicFolderMailboxMailboxStatistics.TotalItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")+[int]$PublicFolderMailboxMailboxStatistics.TotalDeletedItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")
        }
        catch {
           
        }
        
        ##Validate PFMBXsize has excced 80% PublicFolderMailboxSendReceiveQuota
        if ((($PublicFolderMailboxSendReceiveQuota-$PFMBXSizeinB)/(1024*1024*1024)) -le 20) {
            $unhealthyPFMBXcount++
            $unhealthyPFMBX+=$PublicFolderMailbox
        }
        #$percent++
    }
    Write-Host
    write-host "Recommendations:`n================" -ForegroundColor Cyan
    if($unhealthyPFMBXcount -eq 0)
    {
        Write-host "All Public folder mailboxes are on quota healthy state" -ForegroundColor Green
    }
    else {
        Write-host "Please diagnose below public folder mailboxes as their size have exceeded autosplit threshold: " -NoNewline -ForegroundColor Black -BackgroundColor Red
        $unhealthyPFMBX |Format-Table -Wrap -AutoSize Name,Alias,Guid,ExchangeGuid
    }
    #endregion add health quota check on PF MBXs 

    ##Repro that part with smaller values
    #region add health quota check on PFs approaching individual/organization prohibitpostquota,check if we have Giant PFs,that check to be ignored if PFs location is remote
    if ($PublicFoldersLocation -eq "Local") {
        [Int]$unhealthyPFcountapproachingIndQuota=0
        [Int]$unhealthyPFcountapproachingOrgQuota=0
        $unhealthyPF=@()
        $GiantPF=@()
        $UnhealthygiantPF=@()
        foreach($Publicfolder in $Publicfolders)
        {
        [Int64]$OrgPFProhibitPostQuotainB=$OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
        [Int64]$DefaultPublicFolderIssueWarningQuotainB=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[1].split(" ")[0].Replace(",","")
        $Publicfolderstats=Get-PublicFolderStatistics $($Publicfolder.EntryID)
        [Int64]$PublicfolderTotalSize=[Int64]$Publicfolderstats.TotalItemSize.Split("(")[1].split(" ")[0].Replace(",","")+[Int64]$Publicfolderstats.TotalDeletedItemSize.Split("(")[1].split(" ")[0].Replace(",","")
        ##Check health in regards to Organization quota
            if ($Publicfolder.ProhibitPostQuota -eq "unlimited") {
             if ($PublicfolderTotalSize -ge $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -ge 21474836480) {
                $unhealthyPFcountapproachingOrgQuota++
                $UnhealthygiantPF=$UnhealthygiantPF+$publicfolder
                }
                if ($PublicfolderTotalSize -ge $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -le 21474836480) {
                $unhealthyPFcountapproachingOrgQuota++
                $unhealthyPF=$unhealthyPF+$publicfolder
                }
                if ($PublicfolderTotalSize -le $OrgPFProhibitPostQuotainB -and $PublicfolderTotalSize -ge 21474836480) {
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
                    $UnhealthygiantPF=$UnhealthygiantPF+$publicfolder
                }
                 if ($PublicfolderTotalSize -ge $ProhibitPostQuota -and $PublicfolderTotalSize -le 21474836480) {
                    $unhealthyPFcountapproachingIndQuota++
                    $unhealthyPF=$unhealthyPF+$publicfolder
                 }
                if ($PublicfolderTotalSize -le $ProhibitPostQuota -and $PublicfolderTotalSize -ge 21474836480) {
                    $unhealthyPFcountapproachingIndQuota++
                    $GiantPF=$GiantPF+$publicfolder
                }
                #No action done as it's healthy PF
            }
        }
        if ($unhealthyPFcountapproachingOrgQuota -ge 1)
        {   
            Write-host "Please diagnose below public folder(s) as their sizes are not compliant for the below reason: " -ForegroundColor Black -BackgroundColor Red
            if($UnhealthygiantPF.Count -ge 1)
            {
                Write-host "Giant public folder(s) found exceeding OrganizationProhibitPostQuota:`n====================================================================="
                $UnhealthygiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
            if($unhealthyPF.Count -ge 1)
            {
                Write-host "Public folder(s) found exceeding OrganizationProhibitPostQuota:`n==============================================================="
                $unhealthyPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
            if($GiantPF.Count -ge 1)
            {
                Write-host "Giant Public folder(s) found:`n============================="
                $GiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
        }
        if ($unhealthyPFcountapproachingIndQuota -ge 1)
        {
            Write-host "Please diagnose below public folder(s) as their sizes are not compliant for the below reason: "  -ForegroundColor Black -BackgroundColor Red
            if($UnhealthygiantPF.Count -ge 1)
            {
                Write-host "Giant public folder(s) found exceeding ProhibitPostQuota($ProhibitPostQuota Bytes):`n==================================================================================="
                $UnhealthygiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
            if($unhealthyPF.Count -ge 1)
            {
                Write-host "Public folder(s) found exceeding ProhibitPostQuota($ProhibitPostQuota Bytes):`n============================================================================="
                $unhealthyPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
            if($GiantPF.Count -ge 1)
            {
                Write-host "Giant Public folder(s) found:`n============================="
                $GiantPF|Format-Table -Wrap -AutoSize Name,Identity,FolderSize,EntryID
            }
        }
        else 
        {
            Write-Host "All Public folder(s) are on quota healthy state" -ForegroundColor Green
        }

    }
    
    #endregion add health quota check on PFs approaching individual/organization prohibitpostquota,check if we have Giant PFs


    ##add HRR MBXs in case exist
    ##Think about adding Autosplit status
    ##add MEPFs are synced using AD connect
    ##add print for report to html
 }
 
    

<#Function Debug-MEPFNDRCause
{
 Param(
 [parameter(Mandatory=$true)]
 [String]$Cause 
 )

if($Cause -eq "ContentPFMBXfull")
 {
##Validate if Autosplit status is Halted
$PublicFolderMailboxDiagnostics=Get-PublicFolderMailboxDiagnostics $MailPublicFolder.contentmailbox
##Validate Autosplit status
$Autosplitstatus=$PublicFolderMailboxDiagnostics.autosplitinfo.Substring(0,60).split(":")[1].split("")[1]
if($Autosplitstatus -like "Halted")
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
Write-Host "AutoSplit status is Halted so please raise a support request to Microsoft including logs attached under FilePath to solve that issue."
}
elseif($Autosplitstatus -like "SplitCompleted"){
##Validate the date of split
$PublicFolderSplitProcessor=$PublicFolderMailboxDiagnostics.AssistantInfo.ProcessorsState|where {$_ -like "*PublicFolderSplitProcessor*"}
$DateofPublicFolderSplitProcessor=$PublicFolderSplitProcessor.Split("=")[1]
##Validate Autosplitting process was recent
if($DateofPublicFolderSplitProcessor -ge (get-date).AddDays(-7))
{
#Check if DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you might reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.
$DefaultPublicFolderMovedItemRetention=$OrganizationConfig.DefaultPublicFolderMovedItemRetention.Split(":")[0].split(".")[0]
if($DateofPublicFolderSplitProcessor -ge (get-date).AddDays(-$DefaultPublicFolderMovedItemRetention))
{##we might need to lower DefaultPublicFolderMovedItemRetention value to 1 day and invoke mailbox assistant
$UserAction=Read-Host "Organization DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you still need to reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.Do you wish to proceed with that?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention 1.00:00:00
Update-PublicFolderMailbox $MailPublicFolder.contentmailbox
Write-Host "Check later after couple of hours if the $MailPublicFolder.contentmailbox TotalItemSize has reduced by running the below command.`nGet-MailboxStatistics $MailPublicFolder.contentmailbox |ft TotalItemSize `nIf the size is reduced, then the issue is fixed and you may set the MovedItemRetention back to old value of $DefaultPublicFolderMovedItemRetention.00:00:00 using below command.`n Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention $DefaultPublicFolderMovedItemRetention.00:00:00"
}
}
##Something other than DefaultPublicFolderMovedItemRetention value prevented items deletion
else
{
Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
Write-Host "Please raise a support request to Microsoft including logs attached under FilePath to solve that issue."
}
}
##Autosplit was done more than 7 days ago
else
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
Write-Host "Please raise a support request to Microsoft including logs attached under FilePath to solve that issue."
}
}
else
{
##Other Autosplit status 
$PublicFolderSplitProcessor=$PublicFolderMailboxDiagnostics.AssistantInfo.ProcessorsState|where {$_ -like "*PublicFolderSplitProcessor*"}
$DateofPublicFolderSplitProcessor=$PublicFolderSplitProcessor.Split("=")[1]
##Validate Autosplitting process was recent
if($DateofPublicFolderSplitProcessor -ge (get-date).AddDays(-2))
{
#Autosplit process is in PROGRESS, Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
Write-Host "Autosplit process is in PROGRESS, Please raise a support request to Microsoft including logs attached under FilePath to check if there are any issues blocking that progress."
}
##Validate Autosplitting process hasn't ran for more than 2 days
else
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
Write-Host "Please raise a support request to Microsoft including logs attached under FilePath to solve that issue."
}
}
}
if($Cause -eq "OrgProhibitPostQuotaReached")
{
##Increase Org DefaultPublicFolderProhibitPostQuota by 2 GB to mitigate 
##Log the action
$UsernewDefaultPublicFolderProhibitPostQuotavalue=Read-Host "Please insert a new Organization DefaultPublicFolderProhibitPostQuota value in Bytes greater than MEPF size($MEPFTotalSizeinB Bytes)"
$UsernewDefaultPublicFolderIssueWarningQuotavalue=Read-Host "Please insert a new Organization DefaultPublicFolderIssueWarningQuota value in Bytes greater than the old value($DefaultPublicFolderIssueWarningQuotainB Bytes) & lower than $UsernewDefaultPublicFolderProhibitPostQuotavalue Bytes!"
if($UsernewDefaultPublicFolderProhibitPostQuotavalue -ge $DefaultPublicFolderProhibitPostQuotainB -and $UsernewDefaultPublicFolderIssueWarningQuotavalue -ge $DefaultPublicFolderIssueWarningQuotainB -and $UsernewDefaultPublicFolderProhibitPostQuotavalue -ge $UsernewDefaultPublicFolderIssueWarningQuotavalue -and $UsernewDefaultPublicFolderProhibitPostQuotavalue -ge $MEPFTotalSizeinB)
{
Set-OrganizationConfig -DefaultPublicFolderProhibitPostQuota $UsernewDefaultPublicFolderProhibitPostQuotavalue -DefaultPublicFolderIssueWarningQuota $UsernewDefaultPublicFolderIssueWarningQuotavalue
}
else
{
##Values entered are inconsistent
Write-Host "Org. DefaultPublicFolderProhibitPostQuota & DefaultPublicFolderIssueWarningQuota NEW defined values are inconsistent!"
##exit
}
}
if($Cause -eq "IndProhibitPostQuotaReached")
{
$UsernewProhibitPostQuotaValue=Read-Host "Please insert Unlimited to inherit from Organization Public folder Post Quota or a new public folder ProhibitPostQuota value greater than the old value($MEPFProhibitPostQuotainB Bytes) considering that all values entered are in BYTES!"
##Validate warning quota over MEPF
if($MEPFProperties.IssueWarningQuota -like "unlimited")
{
    if($UsernewProhibitPostQuotaValue -like "unlimited")
    {
        try {
            Set-PublicFolder $MEPFProperties.entryid -ProhibitPostQuota $UsernewProhibitPostQuotaValue    
        }
        catch {
            
        }
        
    }
    else {
        if($UsernewProhibitPostQuotaValue -ge $MEPFProhibitPostQuotainB -and $UsernewProhibitPostQuotaValue -ge $MEPFTotalSizeinB)
        {
        Set-PublicFolder $MEPFProperties.entryid -ProhibitPostQuota $UsernewProhibitPostQuotaValue
        }
        else
        {
        write-host "Public folder ProhibitPostQuota NEW defined value is either LOWER than OLD post quota value or PF Item Size value!, please rerun and ensure to specify consistent value!"
        }
        }

    }
    else
        {
        ##IssueWarningQuota has numeric value
        Set-PublicFolder $MEPFProperties.entryid -ProhibitPostQuota $UsernewProhibitPostQuotaValue
        $MEPFIssueWarningQuotainB=$MEPFProperties.IssueWarningQuota.Split("(")[1].split(" ")[0]
        $UsernewIssueWarningQuotaValue=Read-Host "Please insert Unlimited to inherit from Organization Public folder warning Quota or a new public folder IssueWarningQuota value greater than the old value($MEPFIssueWarningQuotainB Bytes) & lower than new Post Quota value $UsernewProhibitPostQuotaValue Bytes! considering that all values entered are in BYTES"
        if($UsernewIssueWarningQuotaValue -le $UsernewProhibitPostQuotaValue -and $UsernewIssueWarningQuotaValue -ge $MEPFIssueWarningQuotainB)
        {
        Set-PublicFolder $MEPFProperties.entryid -IssueWarningQuota $UsernewIssueWarningQuotaValue
        }
        else
        {
        write-host "Public folder IssueWarningQuota NEW defined value is GREATER than NEW defined public folder ProhibitPostQuota value!, please rerun and ensure to specify consistent value!"
        }
}

##Log the action
}
}#>