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
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()

##Code for Menu

$PFMenu=@"
1 - Public folder overview
2 - Diagnosing 554 5.2.2 mailbox full NDR received on sending to MEPF
Q  Quit
     
Select a task by number or Q to quit
"@

$menuchoice=read-host $PFMenu
if ($menuchoice -eq 2)
{
    #region Get the affected MEPF SMTP
    $MEPFSMTP=read-host "Please enter affected MEPF SMTP"
    #endregion Get the affected MEPF SMTP
    Start-MEPFNDRDiagnosis($MEPFSMTP)


$MEPFSMTP="full1@EmbabyTrade.onmicrosoft.com"   
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()
#region Intro with group name 
[string]$SectionTitle = "Introduction"
[String]$article="https://docs.microsoft.com/en-us/exchange/troubleshoot/email-delivery/cannot-send-mail-mepf"
[string]$Description = "This report illustrates causes behind a non-delivery report (NDR) with error code 554 5.2.2 when sending emails to a mail-enabled public folder: "+$MEPFSMTP+", Sections in RED are for checks BLOCKERS while Sections in GREEN are for checks ELIGIBILITIES"
$Description=$Description+",for more informtion please check: $article"
[PSCustomObject]$StartHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate causes in case found by checking FIX section!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
#endregion Intro with group name    
#region global variables used in functions
try {
    [PSCustomObject]$MailPublicFolder=Get-MailPublicFolder $MEPFSMTP
    [PSCustomObject]$ContentPFMBXStatistics=Get-MailboxStatistics $MailPublicFolder.contentmailbox
    [PSCustomObject]$ContentPFMBXProperties=Get-Mailbox -PublicFolder $MailPublicFolder.contentmailbox
    [PSCustomObject]$OrganizationConfig=Get-OrganizationConfig
        
}
catch {
Write-Error -Message "Error"    
}
#endregion global variables used in functions
Start-MEPFNDRDiagnosis($MEPFSMTP)
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
Start-MEPFNDRDiagnosis("full1@EmbabyTrade.onmicrosoft.com")

}

#region Data collection
<#
.SYNOPSIS

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Start-PFDataCollection{

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
   [Int]$PublicFolderMailboxesCount=(Get-Mailbox -PublicFolder -ResultSize unlimited).count
   [Int]$PublicFoldersCount=(Get-PublicFolder -Recurse -ResultSize unlimited).count - 1
   [Int]$MailEnabledPublicFoldersCount=(Get-MailPublicFolder -ResultSize unlimited).count


}

#endregion Data collection




##T.S 554 5.2.2 mailbox full NDR
Function Start-MEPFNDRDiagnosis{
    Param(
        [parameter(Mandatory=$true)]
        [String]$MEPFSMTP)


        $ContentPFMBXProhibitSendReceiveQuotainB=$ContentPFMBXProperties.ProhibitSendReceiveQuota.Split("(")[1].split(" ")[0].Replace(",","")
        [int]$ContentPFMBXSizeinB=[int]$ContentPFMBXStatistics.TotalItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")+[int]$ContentPFMBXStatistics.TotalDeletedItemSize.value.tostring().Split("(")[1].split(" ")[0].Replace(",","")
    
#region Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
if($ContentPFMBXSizeinB -ge $ContentPFMBXProhibitSendReceiveQuotainB)
{
$Orgreached= "Content Public Folder mailbox hosting that mail-enabled public folder has reached its quota!"
#$UserAction=Read-Host "Do you wish to investigate further by checking if Autosplit has processed that mailbox?`nType Y(Yes) to proceed or N(No) to exit!"
<#if ($UserAction -like "*y*")
{
#Call FIX function
Diagnose-MEPFNDRCause("ContentPFMBXfull")
}#>
[string]$SectionTitle = "Validating against content public folder mailbox quota"
[string]$Description = "Checking if the content public folder mailbox hosting the mail-enabled public folder has reached its quota!"
[PSCustomObject]$PFMBXContentQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
$null = $TheObjectToConvertToHTML.Add($PFMBXContentQuotareachedHTML)
Repair-MEPFNDRCause("ContentPFMBXfull")
}
else {
    [string]$SectionTitle = "Validating against content public folder mailbox quota"
    [string]$Description = "Checking if the content public folder mailbox hosting the mail-enabled public folder has reached its quota!"
    [PSCustomObject]$PFMBXContentQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($PFMBXContentQuotareachedHTML)
        
}
#endregion Validating that Content Public Folder mailbox hosting that mail-enabled public folder quota limit is not reached
#region Validating individual/Organization public folder quota

$MEPFStatistics=Get-PublicFolderStatistics -identity $MailPublicFolder.EntryID
$MEPFProperties=Get-PublicFolder $MailPublicFolder.EntryID
[Int]$MEPFTotalSizeinB=$MEPFStatistics.TotalItemSize.Split("(")[1].split(" ")[0].Replace(",","")+$MEPFStatistics.TotalDeletedItemSize.Split("(")[1].split(" ")[0].Replace(",","")
[Int]$MEPFTotalSizeinGB = $MEPFTotalSizeinB /(1024*1024*1024)
##Validate if DefaultPublicFolderProhibitPostQuota at the organization level applies
if($MEPFProperties.ProhibitPostQuota -eq "unlimited")
{
[string]$SectionTitle = "Validating against organization public folder post quota"
[string]$Description = "Checking if public folder total size has reached organization public folder DefaultPublicFolderProhibitPostQuota value!"
##catch unlimited value
$DefaultPublicFolderProhibitPostQuotainB=$OrganizationConfig.DefaultPublicFolderProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
$DefaultPublicFolderIssueWarningQuotainB=$OrganizationConfig.DefaultPublicFolderIssueWarningQuota.Split("(")[1].split(" ")[0].Replace(",","")
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
        [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
        $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
        Repair-MEPFNDRCause("OrgProhibitPostQuotaReached")
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
    [string]$SectionTitle = "Validating against organization public folder post quota"
    [string]$Description = "Checking if public folder total size has reached organization public folder DefaultPublicFolderProhibitPostQuota value!"
    [PSCustomObject]$MEPFOrgPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "No issue found!"
    $null = $TheObjectToConvertToHTML.Add($MEPFOrgPostQuotareachedHTML)
    [string]$SectionTitle = "Validating against individual public folder post quota"
    [string]$Description = "Checking if public folder total size has reached individual public folder ProhibitPostQuota value!"
    [Int]$MEPFProhibitPostQuotainB=$MEPFProperties.ProhibitPostQuota.Split("(")[1].split(" ")[0].Replace(",","")
    [Int]$MEPFProhibitPostQuotainGB = $MEPFProhibitPostQuotainB /(1024*1024*1024)
##Validate that MEPF size is < 20 GB AND greater than Individual ProhibitPostQuota
if($MEPFTotalSizeinB -ge $MEPFProhibitPostQuotainB -and $MEPFTotalSizeinB -le 21474836480)
{
$Orgreached="The individual public folder post quota (ProhibitPostQuota $MEPFProhibitPostQuotainB Bytes) has been reached!`nMail-enabled public folder size ($MEPFTotalSizeinB Bytes) is < 20 GB"
<###Call FIXES function
$UserAction=Read-Host "Do you wish to mitigate the issue by increasing the public folder ProhibitPostQuota value?`nType Y(Yes) to proceed or N(No) to exit!"
if ($UserAction -like "*y*")
{
    Debug-MEPFNDRCause("IndProhibitPostQuotaReached")
}
#>
[PSCustomObject]$PFProhibitPostQuotareachedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring $Orgreached
$null = $TheObjectToConvertToHTML.Add($PFProhibitPostQuotareachedHTML)
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
$Orgreached= "Mail-enabled public folder size ($MEPFTotalSizeinGB GB) is < 20 GB and didn't reach public folder ProhibitPostQuota ($MEPFProhibitPostQuotainGB GB) value!" 
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

Function Repair-MEPFNDRCause
{
 Param(
 [parameter(Mandatory=$true)]
 [String]$Cause 
 )
 if($Cause -eq "OrgProhibitPostQuotaReached")
 {
 [string]$SectionTitle = "FIX"
 $article="https://docs.microsoft.com/en-us/powershell/module/exchange/set-organizationconfig"
 [string]$Description = "Please insert a new Organization DefaultPublicFolderProhibitPostQuota value in correlation with a new DefaultPublicFolderIssueWarningQuota value ensuring that these values are greater than MEPF size($MEPFTotalSizeinB Bytes)using command Set-OrganizationConfig,for more information please check the following article: $article"
 $OrgProhibitQuotaReached="Please ensure to follow the fix to mitigate your issue!"
 [PSCustomObject]$OrgquotaHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $OrgProhibitQuotaReached
 $null = $TheObjectToConvertToHTML.Add($OrgquotaHTML)
 }
 
 if($Cause -eq "IndProhibitPostQuotaReached")
 {
    if($MEPFProperties.IssueWarningQuota -like "unlimited")
    {
#Either increase ProhibitPostQuota by value or set it to unlimited considering is lower than organization configuration value
$valueusedforincrease=$MEPFTotalSizeinB-$MEPFProhibitPostQuotainB

#Please ensure to cover that gap in next defind value

    }
    else {
        
    }




 } 
 if($Cause -eq "ContentPFMBXfull")
 {
##Validate if Autosplit status is Halted
$PublicFolderMailboxDiagnostics=Get-PublicFolderMailboxDiagnostics $MailPublicFolder.contentmailbox
##Validate Autosplit Halted status
$Autosplitstatus=$PublicFolderMailboxDiagnostics.autosplitinfo.Substring(0,60).split(":")[1].split("")[1]
if($Autosplitstatus -like "Halted")
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$SectionTitle = "FIX"
[string]$article="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps"
[string]$Description = "AutoSplit status is Halted so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation,for more information please refer to the following article: $article"
$ContentPFMBXreached="Please ensure to follow the fix to mitigate your issue!"
[PSCustomObject]$AutosplitHaltedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
$null = $TheObjectToConvertToHTML.Add($AutosplitHaltedHTML)
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
if($DateofPublicFolderSplitProcessor -ge (get-date).AddDays(-7))
{
#Check if DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you might reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.
$DefaultPublicFolderMovedItemRetention=$OrganizationConfig.DefaultPublicFolderMovedItemRetention.Split(":")[0].split(".")[0]
if($DateofPublicFolderSplitProcessor -ge (get-date).AddDays(-$DefaultPublicFolderMovedItemRetention))
{
##we might need to lower DefaultPublicFolderMovedItemRetention value to 1 day and invoke mailbox assistant
[string]$SectionTitle = "FIX"
[string]$Description = @"
Organization DefaultPublicFolderMovedItemRetention is keeping the mailbox full, even though AutoSplit completed successfully, you still need to reduce DefaultPublicFolderMovedItemRetention to be 1 day and then invoke mailbox assistant to process the mailbox.Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention 1.00:00:00
Update-PublicFolderMailbox $MailPublicFolder.contentmailbox
Check later after couple of hours if the $MailPublicFolder.contentmailbox TotalItemSize has reduced by running the below command.`nGet-MailboxStatistics $MailPublicFolder.contentmailbox |ft TotalItemSize `nIf the size is reduced, then the issue is fixed and you may set the MovedItemRetention back to old value of $DefaultPublicFolderMovedItemRetention.00:00:00 using below command.`n Set-OrganizationConfig -DefaultPublicFolderMovedItemRetention $DefaultPublicFolderMovedItemRetention.00:00:00
"@
$ContentPFMBXreached="Please ensure to follow the fix to mitigate your issue!"
[PSCustomObject]$AutosplitcompletedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
$null = $TheObjectToConvertToHTML.Add($AutosplitcompletedHTML)
}
}
##Something other than DefaultPublicFolderMovedItemRetention value prevented items deletion
else
{
[string]$SectionTitle = "FIX"
[string]$article="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps"
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation,for more information please refer to the following article: $article"
$ContentPFMBXreached="Please ensure to follow the fix to mitigate your issue!"
[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
##Autosplit was done more than 7 days ago
else
{
#Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$SectionTitle = "FIX"
[string]$article="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps"
[string]$Description = "Please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation,for more information please refer to the following article: $article"
$ContentPFMBXreached="Please ensure to follow the fix to mitigate your issue!"
[PSCustomObject]$AutosplitunknownreasonHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
$null = $TheObjectToConvertToHTML.Add($AutosplitunknownreasonHTML)
}
}
else 
{
##Other Autosplit status 
#Autosplit process is in PROGRESS, Log the PublicFolderMailboxDiagnostics+ContentPFMBXStatistics+ContentPFMBXProperties for customer to raise a support request with it
[string]$SectionTitle = "FIX"
[string]$article="https://docs.microsoft.com/en-us/powershell/module/exchange/get-publicfoldermailboxdiagnostics?view=exchange-ps"
[string]$Description = "Autosplit process is in PROGRESS so please raise a support request to Microsoft including output from Get-PublicFolderMailboxDiagnostics command for further investigation,for more information please refer to the following article: $article"
$ContentPFMBXreached="Please ensure to follow the fix to mitigate your issue!"
[PSCustomObject]$AutosplitHaltedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring $ContentPFMBXreached
$null = $TheObjectToConvertToHTML.Add($AutosplitHaltedHTML)
}
}
##add condition Check prohibitsendquota if it was set to a lower value
}





Function Debug-MEPFNDRCause
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
}

