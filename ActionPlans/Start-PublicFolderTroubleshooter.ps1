##Build a menu for PF T.S Diags starting with 
##     1-PF enviroment general stats (PF Overview,PFs count,MEPFs count,PF MBXs count,DBEB status,Hierarchy sync,autosplit status) 
##     2-T.S 554 5.2.2 mailbox full NDR 
##     3-Diagnose PF dumpster for cases related to either delete items inside PF or PF as a whole
##     4-T.S modern\legacy PFs access in a remote enviroment 
##      
##
##
######################################################################################
##Code for Menu



##T.S 554 5.2.2 mailbox full NDR
#region Get the affected MEPF SMTP
$MEPFSMTP=read-host "Please enter affected MEPF SMTP"
#endregion Get the affected MEPF SMTP
$MailPublicFolder=Get-MailPublicFolder $MEPFSMTP
#region Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
$ContentPFMBXStatistics=Get-MailboxStatistics $MailPublicFolder.contentmailbox
$ContentPFMBXProperties=Get-Mailbox -PublicFolder $MailPublicFolder.contentmailbox
$ContentPFMBXProhibitSendReceiveQuotainGB=$ContentPFMBXProperties.ProhibitSendReceiveQuota.Split("(")[1].split(" ")[0]/(1024*1024*1024)
$ContentPFMBXSizeinGB=$ContentPFMBXStatistics.TotalItemSize.Split("(")[1].split(" ")[0]/(1024*1024*1024)+$ContentPFMBXStatistics.TotalDeletedItemSize.Split("(")[1].split(" ")[0]/(1024*1024*1024)
if($ContentPFMBXSizeinGB -ge $ContentPFMBXProhibitSendReceiveQuotainGB)
{
Write-Host "Content Public Folder mailbox hosting that mail-enabled public folder quota limit is REACHED!"
$UserAction=Read-Host "Do you wish to investigate further by checking if Autosplit has processed that mailbox?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
#Call FIX function
MitigateMEPFNDRCause("ContentPFMBXfull")
}

}
#endregion Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
#region Validating individual/Organization public folder quota
$MEPFStatistics=Get-PublicFolderStatistics $MailPublicFolder.EntryID
$MEPFProperties=Get-PublicFolder $MailPublicFolder.EntryID
##Validate if DefaultPublicFolderProhibitPostQuota at the organization level applies
if($MEPFProperties.ProhibitPostQuota -eq "unlimited")
{
$OrganizationConfig=Get-OrganizationConfig
##catch unlimited value
$DefaultPublicFolderProhibitPostQuotainGB=$OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[1].split(" ")[0]/(1024*1024*1024)
$DefaultPublicFolderIssueWarningQuotainGB=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[1].split(" ")[0]/(1024*1024*1024)
##Test to use foldersize or stick to the below
[Int]$MEPFTotalSizeinGB=$MEPFStatistics.TotalItemSize.Split("(")[1].split(" ")[0]/(1024*1024*1024)+$MEPFTotalItemSize.TotalDeletedItemSize.Split("(")[1].split(" ")[0]/(1024*1024*1024)
##Validate that MEPF size is < 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
if($MEPFTotalSizeinGB-ge $DefaultPublicFolderProhibitPostQuotainGB -and $MEPFTotalSizeinGB -le 20)
{
Write-Host "MEPF size ($MEPFTotalSizeinGB GB) is GREATER THAN Organization DefaultPublicFolderProhibitPostQuota ($DefaultPublicFolderProhibitPostQuotainGB GB)"
###Call FIX function
$UserAction=Read-Host "Do you wish to mitigate the issue by increasing the DefaultPublicFolderProhibitPostQuota & DefaultPublicFolderIssueWarningQuota values?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
MitigateMEPFNDRCause("OrgProhibitPostQuotaReached")
}

}
##Validate that MEPF size is > 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
elseif($MEPFTotalSizeinGB-ge $DefaultPublicFolderProhibitPostQuotainGB -and $MEPFTotalSizeinGB -ge 20)
{
write-host "Mail-enabled public folder size is > 20 GB  AND greater than Organization DefaultPublicFolderProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
}
else
{
write-host "No Issue found.`nMail-enabled public folder size is > 20 GB  AND LOWER than public folder Organzation DefaultPublicFolderProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
}
}
else
{
##Validate that MEPF size is < 20 GB AND greater than Individual ProhibitPostQuota
$MEPFProhibitPostQuotainGB=$MEPFProperties.ProhibitPostQuota.Split("(")[1].split(" ")[0]/(1024*1024*1024)
if($MEPFTotalSizeinGB -ge $MEPFProhibitPostQuotainGB -and $MEPFTotalSizeinGB -le 20)
{
Write-Host "Mail-enabled public folder size ($MEPFTotalSizeinGB GB) is greater than public folder ProhibitPostQuota ($MEPFProhibitPostQuotainGB GB)"
###Call FIX function
Read-Host "Do you wish to mitigate the issue by increasing the public folder ProhibitPostQuota value?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
MitigateMEPFNDRCause("IndProhibitPostQuotaReached")
}
}
##Validate that MEPF size is > 20 GB AND greater than Org DefaultPublicFolderProhibitPostQuota
elseif($MEPFTotalSizeinGB -ge $MEPFProhibitPostQuotainGB -and $MEPFTotalSizeinGB -ge 20)
{
write-host "Mail-enabled public folder size is > 20 GB  AND greater than public folder ProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
}
else
{
write-host "No Issue found.`nMail-enabled public folder size is > 20 GB  AND LOWER than public folder ProhibitPostQuota, we recommend that you delete content from that folder to make it smaller than 20 GB. Or, we recommend that you divide the public folder's content into multiple, smaller public folders as Giant Public Folders impact Autosplitting process!"
}
}

#endregion Validating individual/Organization public folder quota
Function MitigateMEPFNDRCause
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
$UsernewDefaultPublicFolderProhibitPostQuotavalue=Read-Host "Please insert a new Organization DefaultPublicFolderProhibitPostQuota value in GB/MB/KB & greater than the old value($DefaultPublicFolderProhibitPostQuotainGB GB) having the unit(GB/MB/KB) attached to it! ex.3GB(old value) + 2GB=5GB,3MB (old value)+2MB=5MB"
$UsernewDefaultPublicFolderIssueWarningQuotavalue=Read-Host "Please insert a new Organization DefaultPublicFolderIssueWarningQuota value in GB/MB/KB & greater than the old value($DefaultPublicFolderIssueWarningQuotainGB GB) & lower than $UsernewDefaultPublicFolderProhibitPostQuotavalue having the unit(GB/MB/KB) attached to it! ex.2GB(old value) + 1GB=4GB,2MB (old value)+1MB=5MB"
if($UsernewDefaultPublicFolderProhibitPostQuotavalue -ge $DefaultPublicFolderProhibitPostQuotainGB -and $UsernewDefaultPublicFolderIssueWarningQuotavalue -ge $DefaultPublicFolderIssueWarningQuotainGB -and $UsernewDefaultPublicFolderProhibitPostQuotavalue -ge $UsernewDefaultPublicFolderIssueWarningQuotavalue)
{
Set-OrganizationConfig -DefaultPublicFolderProhibitPostQuota $UsernewDefaultPublicFolderProhibitPostQuotavalue -DefaultPublicFolderIssueWarningQuota $UsernewDefaultPublicFolderIssueWarningQuotavalue
}
else
{
##Values entered are inconsistent
Write-Host "Org. DefaultPublicFolderProhibitPostQuota & DefaultPublicFolderIssueWarningQuota NEW defined values are either LOWER than OLD values AND/OR DefaultPublicFolderIssueWarningQuota NEW value is GREATER than DefaultPublicFolderProhibitPostQuota NEW value!"
##exit
}
}
if($Cause -eq "IndProhibitPostQuotaReached")
{
$UsernewProhibitPostQuotaValue=Read-Host "Please insert a new public folder ProhibitPostQuota value in GB/MB/KB & greater than the old value($MEPFProhibitPostQuotainGB GB) having the unit(GB/MB/KB) attached to it! ex.3GB(old value) + 2GB=5GB,3MB (old value)+2MB=5MB"
if($UsernewProhibitPostQuotaValue -ge $MEPFProhibitPostQuotainGB -and $MEPFProperties.IssueWarningQuota -like "unlimited")
{
Set-PublicFolder $MEPFProperties.entryid -ProhibitPostQuota $UsernewProhibitPostQuotaValue
}
elseif($UsernewProhibitPostQuotaValue -le $MEPFProhibitPostQuotainGB)
{
write-host "Public folder ProhibitPostQuota NEW defined value is LOWER than OLD value!, please rerun and ensure to specify a HIGHER value."
}
else
{
##IssueWarningQuota has numeric value
$MEPFIssueWarningQuotainGB=$MEPFProperties.IssueWarningQuota.Split("(")[1].split(" ")[0]/(1024*1024*1024)
$UsernewIssueWarningQuotaValue=Read-Host "Please insert a new public folder IssueWarningQuota value in GB/MB/KB & greater than the old value($MEPFIssueWarningQuotainGB GB) & lower than $UsernewProhibitPostQuotaValue having the unit(GB/MB/KB) attached to it! ex.2GB(old value) + 1GB=3GB,2MB (old value)+1MB=3MB"
if($UsernewIssueWarningQuotaValue -le $UsernewProhibitPostQuotaValue -and $UsernewIssueWarningQuotaValue -ge $MEPFIssueWarningQuotainGB)
{
Set-PublicFolder $MEPFProperties.entryid -IssueWarningQuota $UsernewIssueWarningQuotaValue
}
else
{
write-host "Public folder IssueWarningQuota NEW defined value is either LOWER than OLD IssueWarningQuota value OR GREATER than NEW defined public folder ProhibitPostQuota value!, please rerun and ensure to specify a HIGHER value."
}
}
##Log the action
}
}