<# 1st requirement install the module O365 TS
Import-Module C:\Users\a-haemb\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>
Clear-Host
#$DGConditionsmet=0
#region Connecting to EXO & MSOL

$Workloads = "exo","msol"
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
#endregion Connecting to EXO & MSOL

#region Create working folder for Groups Diag
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\DlToO365GroupUpgradeChecks_$ts"
mkdir $ExportPath -Force |out-null
#endregion Create working folder for Groups Diag

#region Getting the DG SMTP
$dgsmtp=Get-ValidEmailAddress("Email address of the Distribution Group ")
try {
    $dg=get-DistributionGroup -Identity $dgsmtp -ErrorAction stop
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO Directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distribution Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distribution Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
#endregion Getting the DG SMTP
#Array list for collecting all HTML object for creating the report
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()

#region Intro with group name 
[string]$SectionTitle = "Introduction"
[String]$article="https://docs.microsoft.com/en-us/microsoft-365/admin/manage/upgrade-distribution-lists?view=o365-worldwide"
[string]$Description = "This report illustrates Distribution to O365 Group migration eligibility checks taken place over group SMTP: "+$dgsmtp+", Sections in RED are for migration BLOCKERS while Sections in GREEN are for migration ELIGIBILITIES"
$Description=$Description+",for more informtion please check: $article"
[PSCustomObject]$StartHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate migration BLOCKERS in case found!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
#endregion Intro with group name

#Add checking elgibilty for command/ask bhla for running command in whatif upgrade-distributiongroup

#region add check migration is in progress
#endregion 

#Region Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"
$ConditionMemberRestriction=New-Object PSObject
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Join Restriction" -NotePropertyValue $dg.MemberJoinRestriction
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Depart Restriction" -NotePropertyValue $dg.MemberDepartRestriction
[string]$SectionTitle = "Validating Distribution Group Member Restriction Properties"
[string]$Description = "Checking if Distribution Group can't be upgraded if MemberJoinRestriction or MemberDepartRestriction or both values are set to Closed"        
if ($dg.MemberJoinRestriction -eq "Open" -and $dg.MemberDepartRestriction -eq "Open") {
    $ConditionOpenMemberRestriction="Distribution group MemberJoinRestriction & MemberDepartRestriction values are Open"
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionOpenMemberRestriction
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)  
    } 
    else {
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionMemberRestriction -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)
    }
#endRegion Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"

#Region Check if Distribution Group can't be upgraded because it is DirSynced
$ConditionIsDirSynced=New-Object PSObject    
$ConditionIsDirSynced|Add-Member -NotePropertyName "IsDirSynced" -NotePropertyValue $dg.IsDirSynced
[string]$SectionTitle = "Validating Distribution Group IsDirSynced Property"
[string]$Description = "Checking if Distribution Group can't be upgraded because IsDirSynced value is true"    
if ($dg.IsDirSynced -eq $true) {
    [PSCustomObject]$ConditionIsDirSyncedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionIsDirSynced -TableType "List"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)

} 
else {
    [PSCustomObject]$ConditionIsDirSyncedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString "Distribution group is NOT synchronized from On-premises"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)
}
#endRegion Check if Distribution Group can't be upgraded because it is DirSynced

#region Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated
$eap = Get-EmailAddressPolicy -ErrorAction stop
[string]$SectionTitle = "Validating Distribution Group matching EmailAddressPolicy"
[string]$Description = "Checking if Distribution Group can't be upgraded because Admin has applied Group Email Address Policy for the groups on the organization"
$ConditionEAP=New-Object PSObject    
# Bypass that step if there's no EAP 
 if($null -ne $eap)
 {
 $matchingEap = @( $eap | where-object{$_.RecipientFilter -eq "RecipientTypeDetails -eq 'GroupMailbox'" -and $_.EnabledEmailAddressTemplates.ToString().Split("@")[1] -ne $dg.PrimarySmtpAddress.ToString().Split("@")[1]})
 if ($matchingEap.Count -ne 0) {
     $count=1
     foreach($matcheap in $matchingEap)
     {
         $ConditionEAP|Add-Member -NotePropertyName "EmailAddressPolicy$count Name" -NotePropertyValue $matcheap
         $count++
    }
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionEAP -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
    
}
else {
    $ConditionNOEAP="NO matching Group Email Address Policy for the groups on the organization"
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionNOEAP
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}
 }
else {
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataString $ConditionNOEAP
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}

 #endregion Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated
<#
#region Check if Distribution Group can't be upgraded because DlHasChildGroups
[string]$SectionTitle = "Validating Distribution Group Child Membership"
[string]$Description = "Checking if Distribution Group can't be upgraded because it contains child groups"
$ConditionChildDG=New-Object PSObject    
try {
    $members = Get-DistributionGroupMember $($dg.Guid.ToString()) -ErrorAction stop
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
$childgroups = $members | Where-Object{ $_.RecipientTypeDetails -eq "MailUniversalDistributionGroup"}
if ($null -ne $childgroups) {
    $count=1
    foreach($childgroup in $childgroups)
    {
        $ConditionChildDG|Add-Member -NotePropertyName "Child Group$count ALias" -NotePropertyValue $childgroup.Alias
        $count++
    }
    [PSCustomObject]$ConditionChildDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionChildDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionChildDGHTML)
} 
else {
    $ConditionChildDG|Add-Member -NotePropertyName "Child Group ALias" -NotePropertyValue "No child groups found"
    [PSCustomObject]$ConditionChildDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionChildDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionChildDGHTML)
}
#endregion Check if Distribution Group can't be upgraded because DlHasChildGroups
#>
#region Check if Distribution Group can't be upgraded because DlHasParentGroups
[string]$SectionTitle = "Validating Distribution Group Parent Membership"
[string]$Description = "Checking if Distribution Group can't be upgraded because it is a child group of another parent group"
$ConditionParentDG=@()
try {
    $alldgs=Get-DistributionGroup -ResultSize unlimited -ErrorAction Stop
    $CurrentProperty = "Retrieving All DGs in the EXO directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve All DGs" -Step $CurrentProperty -Description $CurrentDescription
    
}
catch {
    $CurrentProperty = "Retrieving All DGs in the EXO directory"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve All DGs" -Step $CurrentProperty -Description $CurrentDescription
}  
$parentdgcount=1
foreach($parentdg in $alldgs)
{
    try {
        $Pmembers = Get-DistributionGroupMember $($parentdg.Guid.ToString()) -ErrorAction Stop
        $CurrentProperty = "Retrieving: $parentdg members"
        $CurrentDescription = "Success"
        write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
    }
    catch {
        $CurrentProperty = "Retrieving: $parentdg members"
        $CurrentDescription = "Failure"
        write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
    }

foreach ($member in $Pmembers)
{if ($member.alias -like $dg.alias)
{
    $ConditionParentDG+=$parentdg
    $parentdgcount++
}
}
}
if($parentdgcount -le 1)
{
    [String]$NoParentDG="Distribution group is NOT a member of another group"
    [PSCustomObject]$ConditionParentDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $NoParentDG
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
else {
    $ConditionParentDG = $ConditionParentDG |Select-Object @{ Name = 'Parent Group Display Name';  Expression = {$_.DisplayName}},@{ Name = 'Parent Group Alias';  Expression = {$_.Alias}},@{ Name = 'Parent Group GUID';  Expression = {$_.GUID}},@{ Name = 'Parent Group RecipientTypeDetails';  Expression = {$_.RecipientTypeDetails}},@{ Name = 'Parent Group PrimarySmtpAddress';  Expression = {$_.PrimarySmtpAddress}}
    [PSCustomObject]$ConditionParentDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionParentDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
#endregion Check if Distribution Group can't be upgraded because DlHasParentGroups

#region Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes with RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser
[string]$SectionTitle = "Validating Distribution Group Members Recipient Types"
[string]$Description = "Checking if Distribution Group can't be upgraded because DL contains member RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser"
#$ConditionDGmembers=New-Object psobject
try {
    $members = Get-DistributionGroupMember $($dg.Guid.ToString()) -ErrorAction stop
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
$matchingMbr = @( $members | Where-Object {$_.RecipientTypeDetails -ne "UserMailbox" -and `
        $_.RecipientTypeDetails -ne "SharedMailbox" -and `
        $_.RecipientTypeDetails -ne "TeamMailbox" -and `
        $_.RecipientTypeDetails -ne "MailUser" -and `
        $_.RecipientTypeDetails -ne "GuestMailUser" -and `
        $_.RecipientTypeDetails -ne "RoomMailbox" -and `
        $_.RecipientTypeDetails -ne "EquipmentMailbox" -and `
        $_.RecipientTypeDetails -ne "User" -and `
        $_.RecipientTypeDetails -ne "DisabledUser" `
        
})

if($matchingMbr.Count -ge 1)
{
    $matchingMbr=$matchingMbr|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
    [PSCustomObject]$ConditionDGmembersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $matchingMbr -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)

    } 
else {
    [String]$ConditionDGmembers="Distribution group contains supported members"
    [PSCustomObject]$ConditionDGmembersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionDGmembers
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)
    }

#endregion Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes with RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser

#region Check if Distribution Group can't be upgraded because it has more than 100 owners or it has no owner
[string]$SectionTitle = "Validating Distribution Group Owners Count"
[string]$Description = "Checking if Distribution Group can't be upgraded because it has more than 100 owners or it has no owners"
$ConditionDGowners=New-Object PSObject
$owners=$dg.ManagedBy
if ($owners.Count -gt 100) {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "Owners are greater than 100"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
} 
if ($owners.Count -eq 0) {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "No owners found"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
else {
    $DGownersfound="Distrubtion group Owners found and are less than 100"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $DGownersfound
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution list which has more than 100 owners or it has no owner

#region Check if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL
[string]$SectionTitle = "Validating Distribution Group Sender Restriction"
[string]$Description = "Checking if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL"
$ConditionDGSender=@()
[int]$SenderRestrictionCount=1
foreach($alldg in $alldgs)
{
if ($alldg.AcceptMessagesOnlyFromSendersOrMembers -like $dg.Name -or $alldg.AcceptMessagesOnlyFromDLMembers -like $dg.Name )
{
    
    $ConditionDGSender=$ConditionDGSender+$alldg
    $SenderRestrictionCount++
}
}
if ($SenderRestrictionCount -le 1) {
    $NoDGSenderfound="Distribution group is NOT part of Sender Restriction in another group"
    [PSCustomObject]$ConditionDGSenderHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $NoDGSenderfound
    $null = $TheObjectToConvertToHTML.Add($ConditionDGSenderHTML)
}
else {
    $ConditionDGSender=$ConditionDGSender|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
    [PSCustomObject]$ConditionDGSenderHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGSender -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGSenderHTML)
}

#endregion Check if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL

#region Check if Distribution Group can't be upgraded because Distribution lists which were converted to RoomLists or isn't a security group nor Dynamic DG
$Conditionnonsupportedrec=New-Object PSObject
[string]$SectionTitle = "Validating Distribution Group RecipientTypeDetails Property"
[string]$Description = "Checking if Distribution Group can't be upgraded because it was converted to RoomList or isn't a security group nor Dynamic DG"
$Conditionnonsupportedrec|Add-Member -NotePropertyName "Group RecipientTypeDetails" -NotePropertyValue $dg.RecipientTypeDetails
if($dg.RecipientTypeDetails -like "MailUniversalSecurityGroup" -or $dg.RecipientTypeDetails -like "DynamicDistributionGroup" -or $dg.RecipientTypeDetails -like "roomlist" ) 
{
    [PSCustomObject]$ConditionnonsupportedrecHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditionnonsupportedrec -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
else {
    $supportedrec="Distribution group isn't a Security group nor a Dynamic distribution group nor converted to a RoomList"
    [PSCustomObject]$ConditionnonsupportedrecHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $supportedrec
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution lists which were converted to RoomLists or isn't a security group nor Dynamic DG

#region Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox
$Conditionfwdmbx=@()
[string]$SectionTitle = "Validating Distribution Group Forwarding Usage"
[string]$Description = "Checking if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox"

try {
    $sharedMBXs=Get-Mailbox -ResultSize unlimited -RecipientTypeDetails sharedmailbox -ErrorAction stop
    $CurrentProperty = "Retrieving All Shared MBXs in the EXO directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Shared Mailboxes" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving All Shared MBXs in the EXO directory"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Shared Mailboxes" -Step $CurrentProperty -Description $CurrentDescription
}
$counter=1
foreach($sharedMBX in $sharedMBXs)
{
    if ($sharedMBX.ForwardingAddress -match $dg.name -or $sharedMBX.ForwardingSmtpAddress -match $dg.PrimarySmtpAddress)
    {
        $Conditionfwdmbx= $Conditionfwdmbx+$sharedMBX
        $counter++
    }
}
if ($counter -le 1) {
    $Nofwdmbxfound="Distribution group is NOT configured to be a forwarding address for any Shared Mailbox"
    [PSCustomObject]$ConditionfwdmbxHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $Nofwdmbxfound
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
else {
    $Conditionfwdmbx=$Conditionfwdmbx|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
    [PSCustomObject]$ConditionfwdmbxHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditionfwdmbx -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
#endregion Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox

#region Check for duplicate Alias,PrimarySmtpAddress,Name,DisplayName on EXO objects
$Conditiondupobj=@()
[string]$SectionTitle = "Validating Distribution Group Duplicates"
[string]$Description = "Checking if Distribution Group can't be upgraded because duplicate objects having same Alias,PrimarySmtpAddress,Name,DisplayName found"
try {
    $dupAlias=Get-Recipient -IncludeSoftDeletedRecipients -Identity $dg.alias -ResultSize unlimited -ErrorAction stop
    $dupAddress=Get-Recipient -IncludeSoftDeletedRecipients -ResultSize unlimited -Identity $dg.PrimarySmtpAddress -ErrorAction stop
    $dupDisplayName=Get-Recipient -IncludeSoftDeletedRecipients -ResultSize unlimited -Identity $dg.DisplayName -ErrorAction stop
    $dupName=Get-Recipient -IncludeSoftDeletedRecipients -ResultSize unlimited -Identity $dg.Name -ErrorAction stop
    $CurrentProperty = "Retrieving duplicate recipients having same Alias,PrimarySmtpAddress,Name,DisplayName in the EXO directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Duplicate Recipient Objects" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving duplicate recipients having same Alias,PrimarySmtpAddress,Name,DisplayName in the EXO directory"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Duplicate Recipient Objects" -Step $CurrentProperty -Description $CurrentDescription
    
}
    if($dupAlias.Count -ge 2 -or $dupAddress.Count -ge 2 -or $dupDisplayName.Count -ge 2 -or $dupName.Count -ge 2)
    {
        if($dupAlias.Count -ge 2)
        {   $dupalias=$dupalias|where-object {$_.guid -notlike $dg.guid}
            $Conditiondupobj=$Conditiondupobj+$dupalias

        }
        elseif ($dupAddress.Count -ge 2) {
            $dupAddress=$dupAddress|where-object {$_.guid -notlike $dg.guid}
            $Conditiondupobj=$Conditiondupobj+$dupAddress
        }
        elseif ($dupDisplayName.Count -ge 2) {
            $dupDisplayName=$dupDisplayName|where-object {$_.guid -notlike $dg.guid}
            $Conditiondupobj=$Conditiondupobj+$dupDisplayName
        }
        elseif ($dupName.Count -ge 2) {
            $dupName=$dupName|where-object {$_.guid -notlike $dg.guid}
            $Conditiondupobj=$Conditiondupobj+$dupName
        }
        $Conditiondupobj=$Conditiondupobj|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
        [PSCustomObject]$ConditiondupobjHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditiondupobj -TableType "Table"
        $null = $TheObjectToConvertToHTML.Add($ConditiondupobjHTML)
    }
    else {
        $Nodupobjfound="No Duplicate objects found sharing same Alias,PrimarySmtpAddress,Name & DisplayName"
        [PSCustomObject]$ConditiondupobjHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $Nodupobjfound
        $null = $TheObjectToConvertToHTML.Add($ConditiondupobjHTML)    
    }


#endregion Check for duplicate Alias,PrimarySmtpAddress,Name,DisplayName on EXO objects


##Repro is done for all except EAP condition

<#region finalizescript--Pending
if($DGConditionsmet -gt 0){
    "DG Upgrade Failed"|Out-file -FilePath $ExportPath\result.txt

}
else{
    "DG Upgrade Succeeded"|Out-file -FilePath $ExportPath\result.txt
    
}
#>

#region ResultReport
[string]$FilePath = $ExportPath + "\DistributionGroupUpgradeCheck.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Distribution Group Upgrade Checker" -ReportTitle "Distribution Group Upgrade Checker" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
#Question to ask enduser for opening the HTMl report
$OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile -like "*y*")
{
Write-Host "Opening report...." -ForegroundColor Cyan
Start-Process $FilePath
}
#endregion ResultReport
   
# End of the Diag
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu
