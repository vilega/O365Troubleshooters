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
    write-log -Function "Retrieve Distrubtion Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distrubtion Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
#endregion Getting the DG SMTP

#Region Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()
$ConditionMemberRestriction=New-Object PSObject
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Join Restriction" -NotePropertyValue $dg.MemberJoinRestriction
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Depart Restriction" -NotePropertyValue $dg.MemberDepartRestriction
[string]$SectionTitle = "Validating Distribution Group Member Restriction Properties"
[string]$Description = "Checking if Distribution Group can't be upgraded if MemberJoinRestriction or MemberDepartRestriction or both values are set to Closed"        
if ($dg.MemberJoinRestriction -eq "Open" -and $dg.MemberDepartRestriction -eq "Open") {
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionMemberRestriction -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)  
    } 
    else {
    $DGConditionsmet++    
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionMemberRestriction -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)
    }
#endRegion Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"

#Region Check if Distribution Group can't be upgraded because it is DirSynced
$ConditionIsDirSynced=New-Object PSObject    
$ConditionIsDirSynced|Add-Member -NotePropertyName "IsDirSynced" -NotePropertyValue $dg.IsDirSynced
$ConditionIsDirSynced|Add-Member -NotePropertyName "HtmlIssue" -NotePropertyValue "PLZ Ignore"
[string]$SectionTitle = "Validating Distribution Group IsDirSynced Property"
[string]$Description = "Checking if Distribution Group can't be upgraded because IsDirSynced value is true"    
if ($dg.IsDirSynced -eq $true) {
    $DGConditionsmet++ 
    [PSCustomObject]$ConditionIsDirSyncedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionIsDirSynced -TableType "List"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)

} 
else {
    [PSCustomObject]$ConditionIsDirSyncedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionIsDirSynced -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)
}
#endRegion Check if Distribution Group can't be upgraded because it is DirSynced

#region Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated
$eap = Get-EmailAddressPolicy -ErrorAction stop
[string]$SectionTitle = "Validating Distribution Group matching EmailAddressPolicy"
[string]$Description = "Checking if Distribution Group can't be upgraded because of matching EmailAddressPolicy"
$ConditionEAP=New-Object PSObject    
# Bypass that step if there's no EAP 
 if($eap -ne $null)
 {
 $matchingEap = @( $eap | where-object{$_.RecipientFilter -eq "RecipientTypeDetails -eq 'GroupMailbox'" -and $_.EnabledEmailAddressTemplates.AddressTemplateString.ToString().Split("@")[1] -ne $dg.PrimarySmtpAddress.Domain.ToString()} )
 if ($matchingEap.Count -ne 0) {
     $count=1
     foreach($matcheap in $matchingEap)
     {
         $ConditionEAP|Add-Member -NotePropertyName "EmailAddressPolicy$count Name" -NotePropertyValue $matcheap
         $count++
    }
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionEAP -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
    
}
else {
    $ConditionEAP|Add-Member -NotePropertyName "EmailAddressPolicy Name" -NotePropertyValue "No matching EmailAddressPolicy"
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionEAP -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}
 }
else {
    $ConditionEAP|Add-Member -NotePropertyName "EmailAddressPolicy" -NotePropertyValue "No matching EmailAddressPolicy"
    $ConditionEAP|Add-Member -NotePropertyName "EmailAddressPolicy1" -NotePropertyValue "aaa"
    [PSCustomObject]$ConditionEAPHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionEAP -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}

 #endregion Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated

#region Check if Distribution Group can't be upgraded because DlHasChildGroups
[string]$SectionTitle = "Validating Distribution Group Child Membership"
[string]$Description = "Checking if Distribution Group can't be upgraded because it contains child groups"
$ConditionChildDG=New-Object PSObject    
try {
    $members = Get-DistributionGroupMember $($dg.Guid.ToString()) -ErrorAction stop
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distrubtion Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp members"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distrubtion Group membership" -Step $CurrentProperty -Description $CurrentDescription
}
$childgroups = $members | Where-Object{ $_.RecipientTypeDetails -eq "MailUniversalDistributionGroup"}
if ($childgroups -ne $null) {
    $count=1
    foreach($childgroup in $childgroups)
    {
        $ConditionChildDG|Add-Member -NotePropertyName "Child Group$count ALias" -NotePropertyValue $childgroup.Alias
        $count++
    }
    [PSCustomObject]$ConditionChildDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionChildDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionChildDGHTML)
} 
else {
    $ConditionChildDG|Add-Member -NotePropertyName "Child Group ALias" -NotePropertyValue "No child groups found"
    $ConditionChildDG|Add-Member -NotePropertyName "Html issue" -NotePropertyValue "Test"
    [PSCustomObject]$ConditionChildDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionChildDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionChildDGHTML)
}
#endregion Check if Distribution Group can't be upgraded because DlHasChildGroups

#region Check if Distribution Group can't be upgraded because DlHasParentGroups
[string]$SectionTitle = "Validating Distribution Group Parent Membership"
[string]$Description = "Checking if Distribution Group can't be upgraded because it is a child group of another parent group"
$ConditionParentDG=New-Object PSObject  
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
        write-log -Function "Retrieve Distrubtion Group membership" -Step $CurrentProperty -Description $CurrentDescription
    }
    catch {
        $CurrentProperty = "Retrieving: $parentdg members"
        $CurrentDescription = "Failure"
        write-log -Function "Retrieve Distrubtion Group membership" -Step $CurrentProperty -Description $CurrentDescription
    }

foreach ($member in $Pmembers)
{if ($member.alias -like $dg.alias)
{
    $ConditionParentDG|Add-Member -NotePropertyName "Parent Group$parentdgcount Alias" -NotePropertyValue $parentdg.Alias
    $parentdgcount++
}
}
}
if($parentdgcount -le 1)
{
    $ConditionParentDG|Add-Member -NotePropertyName "Parent Group Alias" -NotePropertyValue "No parent groups found"
    [PSCustomObject]$ConditionParentDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionParentDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
else {
    [PSCustomObject]$ConditionParentDGHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionParentDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
#endregion Check if Distribution Group can't be upgraded because DlHasParentGroups

#region Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes
[string]$SectionTitle = "Validating Distribution Group Members Recipient Types"
[string]$Description = "Checking if Distribution Group can't be upgraded because of unsupported member types"
$ConditionDGmembers=New-Object PSObject
$matchingMbr = @( $members | Where-Object { $_.RecipientTypeDetails -ne "UserMailbox" -and `
        $_.RecipientTypeDetails -ne "SharedMailbox" -and `
        $_.RecipientTypeDetails -ne "TeamMailbox" -and `
        $_.RecipientTypeDetails -ne "MailUser" -and `
        $_.RecipientTypeDetails -ne "GuestMailUser" -and `
        $_.RecipientTypeDetails -ne "RoomMailbox" -and `
        $_.RecipientTypeDetails -ne "EquipmentMailbox" -and `
        $_.RecipientTypeDetails -ne "User" -and `
        $_.RecipientTypeDetails -ne "DisabledUser" `
})
if ($matchingMbr.Count -ne 0) {
    $counter=1
    foreach($matchedmbr in $matchingMbr)
    {
        $ConditionDGmembers|Add-Member -NotePropertyName "Unsuppported member$counter" -NotePropertyValue $matchedmbr.Alias
        $counter++
    }
    [PSCustomObject]$ConditionDGmembersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGmembers -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)

    } 
else {
    $ConditionDGmembers|Add-Member -NotePropertyName "Unsuppported members Count" -NotePropertyValue "No unsupported members found"
    [PSCustomObject]$ConditionDGmembersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGmembers -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)
    }
#endregion Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes

#region Check if Distribution Group can't be upgraded because it has more than 100 owners or it has no owner
[string]$SectionTitle = "Validating Distribution Group Owners Count"
[string]$Description = "Checking if Distribution Group can't be upgraded because it has more than 100 owners or it has no owners"
$ConditionDGowners=New-Object PSObject
$ConditionDGowners|Add-Member -NotePropertyName "HTMLIssue" -NotePropertyValue "test"
$owners=$dg.ManagedBy
if ($owners.Count -gt 100) {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "Owners are greater than 100"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
} 
if ($owners.Count -eq 0) {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "No owners found"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
else {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "Owners are less than 100"
    [PSCustomObject]$ConditionDGownersHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution list which has more than 100 owners or it has no owner

#region Check if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL
[string]$SectionTitle = "Validating Distribution Group Sender Restriction"
[string]$Description = "Checking if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL"
$ConditionDGSender=New-Object PSObject
[int]$SenderRestrictionCount=1
foreach($alldg in $alldgs)
{
if ($alldg.AcceptMessagesOnlyFromSendersOrMembers -match $dg.Alias -or $alldg.AcceptMessagesOnlyFromDLMembers -match $dg.Alias )
{
    
    $ConditionDGSender|Add-Member -NotePropertyName "Group ALias$SenderRestrictionCount Having Sender Restriction" -NotePropertyValue $alldg.Alias
    $SenderRestrictionCount++
}
}
if ($SenderRestrictionCount -le 1) {
    $ConditionDGSender|Add-Member -NotePropertyName "Group ALias Having Sender Restriction" -NotePropertyValue "No sender restrictions found"
    [PSCustomObject]$ConditionDGSenderHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGSender -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGSenderHTML)
}
else {
    [PSCustomObject]$ConditionDGSenderHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionDGSender -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGSenderHTML)
}

#endregion Check if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL

#region Check if Distribution Group can't be upgraded because Distribution lists which were converted to RoomLists or isn't a security group nor Dynmaic DG
$Conditionnonsupportedrec=New-Object PSObject
[string]$SectionTitle = "Validating Distribution Group RecipientTypeDetails Property"
[string]$Description = "Checking if Distribution Group can't be upgraded because it was converted to RoomList or isn't a security group nor Dynmaic DG"
$Conditionnonsupportedrec|Add-Member -NotePropertyName "Group RecipientTypeDetails" -NotePropertyValue $dg.RecipientTypeDetails
$Conditionnonsupportedrec|Add-Member -NotePropertyName "HTMLIssue" -NotePropertyValue "Test"
if($dg.RecipientTypeDetails -like "MailUniversalSecurityGroup" -or $dg.RecipientTypeDetails -like "DynamicDistributionGroup" -or $dg.RecipientTypeDetails -like "roomlist" ) 
{
    [PSCustomObject]$ConditionnonsupportedrecHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditionnonsupportedrec -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
else {
    [PSCustomObject]$ConditionnonsupportedrecHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditionnonsupportedrec -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution lists which were converted to RoomLists or isn't a security group nor Dynmaic DG

#region Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox
$Conditionfwdmbx=New-Object PSObject
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
    if ($sharedMBX.ForwardingAddress -match $dg.alias -or $sharedMBX.ForwardingSmtpAddress -match $dg.alias)
    {
        $Conditionfwdmbx|Add-Member -NotePropertyName "Shared Mailbox$counter Alias" -NotePropertyValue $sharedMBX.Alias
        $counter++
    }
}
if ($counter -le 1) {
    $Conditionfwdmbx|Add-Member -NotePropertyName "Shared Mailbox Alias" -NotePropertyValue "No shared MBX has that DL address configured for forwarding"
    [PSCustomObject]$ConditionfwdmbxHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditionfwdmbx -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
else {
    [PSCustomObject]$ConditionfwdmbxHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditionfwdmbx -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
#endregion Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox

#region Check for duplicate Alias,PrimarySmtpAddress,Name,DisplayName on EXO objects
$Conditiondupobj=New-Object PSObject
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
    {$counter=1
        if($dupAlias.Count -ge 2)
        {
            foreach($ObjectAlias in $dupAlias)
            {   if ($ObjectAlias.Guid -notmatch $dg.Guid) {
                $Conditiondupobj|Add-Member -NotePropertyName "Duplicate Object$counter" -NotePropertyValue $ObjectAlias.PrimarySmtpAddress
                $counter++
            }
            }
        }
        elseif ($dupAddress.Count -ge 2) {
            foreach($ObjectAlias in $dupAddress)
            {   if ($ObjectAlias.Guid -notmatch $dg.Guid) {
                $Conditiondupobj|Add-Member -NotePropertyName "Duplicate Object$counter" -NotePropertyValue $ObjectAlias.PrimarySmtpAddress
                $counter++
            }   
                
            }
        }
        elseif ($dupDisplayName.Count -ge 2) {
            foreach($ObjectAlias in $dupDisplayName)
            {   if ($ObjectAlias.Guid -notmatch $dg.Guid) {
                $Conditiondupobj|Add-Member -NotePropertyName "Duplicate Object$counter" -NotePropertyValue $ObjectAlias.PrimarySmtpAddress
                $counter++
            }
            }
        }
        elseif ($dupName.Count -ge 2) {
            foreach($ObjectAlias in $dupName)
            {   if ($ObjectAlias.Guid -notmatch $dg.Guid) {
                $Conditiondupobj|Add-Member -NotePropertyName "Duplicate Object$counter" -NotePropertyValue $ObjectAlias.PrimarySmtpAddress
                $counter++
            }
            }
        }
        [PSCustomObject]$ConditiondupobjHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditiondupobj -TableType "Table"
        $null = $TheObjectToConvertToHTML.Add($ConditiondupobjHTML)
    }
    else {
        $Conditiondupobj|Add-Member -NotePropertyName "Duplicate Object" -NotePropertyValue "No duplicate objects found"
        [PSCustomObject]$ConditiondupobjHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $Conditiondupobj -TableType "Table"
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
[string]$FilePath = $ExportPath + "\DistrubtionGroupUpgradeCheck.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Distrubtion Group Upgrade Checker" -ReportTitle "Distrubtion Group Upgrade Checker" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
##Pending question to ask enduser for opening the HTMl report
#endregion ResultReport


    
# End of the Diag
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu
