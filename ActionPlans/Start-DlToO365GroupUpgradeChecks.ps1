<# 1st requirement install the module O365 TS
Import-Module C:\Users\haembab\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>

<#
.Synopsis
   Checks for blockers preventing Distribution Group migration.
.DESCRIPTION
   This script illustrates Distribution to M365 Group migration eligibility checks taken place over the provided distribution group, most of the checks are listed on the below article!
   https://docs.microsoft.com/en-us/microsoft-365/admin/manage/upgrade-distribution-lists?view=o365-worldwide
   For more information on the script procedures please check the below article!
   https://aka.ms/O365Troubleshooters/DlToGroupUpgrade
.INPUTS
  None.
.OUTPUTS
   HTML report showing distribution group migration blockers & eligibilities and log file
.LINK
   https://aka.ms/O365Troubleshooters/DlToGroupUpgrade
#>

Clear-Host

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
Clear-Host
#region Create working folder for Groups Diag
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\DlToO365GroupUpgradeChecks_$ts"
mkdir $ExportPath -Force |out-null
#endregion Create working folder for Groups Diag

#region Getting the DG SMTP
$dgsmtp=Get-ValidEmailAddress("Email address of the Distribution Group ")
try {
    $dg=get-DistributionGroup -Identity $dgsmtp -ErrorAction stop -ResultSize unlimited
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO Directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distribution Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distribution Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
    $Errorencountered=$Global:error[0].Exception
    Write-Host "Error encountered during executing the script!"-ForegroundColor Red
    Write-Host $Errorencountered -ForegroundColor Red
    Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
    Start-Sleep -Seconds 3
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
    ##write log and exit function
}
#endregion Getting the DG SMTP
#Array list for collecting all HTML object for creating the report
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()

#region Intro with group name 
$blockersinhtml='<span style="color: red">BLOCKERS</span>'
$Eligibilitiesinhtml='<span style="color: green">ELIGIBILITIES</span>'
$Greeninhtml='<span style="color: green">GREEN</span>'
$Redinhtml='<span style="color: red">RED</span>'
[string]$SectionTitle = "Introduction"
[String]$article='<a href="https://docs.microsoft.com/en-us/microsoft-365/admin/manage/upgrade-distribution-lists?view=o365-worldwide" target="_blank">Upgrade distribution lists to Microsoft 365 Groups in Outlook</a>'
[string]$Description = "This report illustrates Distribution to M365 Group migration eligibility checks taken place over group SMTP: "+"<b>$dgsmtp</b>"+", Sections in$Redinhtml are for migration$blockersinhtml while Sections in$Greeninhtml are for migration$Eligibilitiesinhtml"
$Description=$Description+",for more informtion please check: $article"
[PSCustomObject]$StartHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate $blockersinhtml in case found!"
#[PSCustomObject]$StartHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please ensure to mitigate migration BLOCKERS in case found!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
#endregion Intro with group name


#Region Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"
$ConditionMemberRestriction=New-Object PSObject
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Join Restriction" -NotePropertyValue $dg.MemberJoinRestriction
$ConditionMemberRestriction|Add-Member -NotePropertyName "Member Depart Restriction" -NotePropertyValue $dg.MemberDepartRestriction
[string]$SectionTitle = "Validating Distribution Group Member Restriction Properties"
[string]$Description = "Checking if Distribution Group can't be upgraded if MemberJoinRestriction or MemberDepartRestriction or both values are set to Closed"        
if ($dg.MemberJoinRestriction -eq "Open" -and $dg.MemberDepartRestriction -eq "Open") {
    $ConditionOpenMemberRestriction="Distribution group MemberJoinRestriction & MemberDepartRestriction values are Open"
    [PSCustomObject]$ConditionMemberRestrictionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionOpenMemberRestriction
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)  
    } 
    else {
    [PSCustomObject]$ConditionMemberRestrictionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionMemberRestriction -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)
    }
#endRegion Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"

#Region Check if Distribution Group can't be upgraded because it is DirSynced
$ConditionIsDirSynced=New-Object PSObject    
$ConditionIsDirSynced|Add-Member -NotePropertyName "IsDirSynced" -NotePropertyValue $dg.IsDirSynced
[string]$SectionTitle = "Validating Distribution Group IsDirSynced Property"
[string]$Description = "Checking if Distribution Group can't be upgraded because IsDirSynced value is true"    
if ($dg.IsDirSynced -eq $true) {
    [PSCustomObject]$ConditionIsDirSyncedHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionIsDirSynced -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)

} 
else {
    [PSCustomObject]$ConditionIsDirSyncedHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString "Distribution group is NOT synchronized from On-premises"
    $null = $TheObjectToConvertToHTML.Add($ConditionIsDirSyncedHTML)
}
#endRegion Check if Distribution Group can't be upgraded because it is DirSynced

#region Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated
$eap = Get-EmailAddressPolicy -ErrorAction stop
[string]$SectionTitle = "Validating Distribution Group matching EmailAddressPolicy"
[string]$Description = "Checking if Distribution Group can't be upgraded because Admin has applied Group Email Address Policy for the groups on the organization e.g. DL PrimarySmtpAddress @"+"<b>C</b>"+"ontoso.com while the EAP EnabledPrimarySMTPAddressTemplate is @"+"<b>c</b>"+"ontoso.com (case-sensitive condition should match) OR DL PrimarySmtpAddress @contoso.com however there's an EAP with EnabledPrimarySMTPAddressTemplate set to @fabrikam.com"
$ConditionEAP=New-Object PSObject    
# Bypass that step if there's no EAP 
 if($null -ne $eap)
 {
     #Case sensitive operator to catch any difference even in letters of smtp address

 $ViolatedEap = @( $eap | where-object{$_.RecipientFilter -eq "RecipientTypeDetails -eq 'GroupMailbox'" -and $_.EnabledPrimarySMTPAddressTemplate.ToString().Split("@")[1] -cne $dg.PrimarySmtpAddress.ToString().Split("@")[1]})
 if ($ViolatedEap.Count -ge 1) {
    $ConditionEAP=$ViolatedEap|Select-Object Identity,Priority,@{label='PrimarySMTPAddressTemplate';expression={($_.EnabledPrimarySMTPAddressTemplate).split("@")[1]}} |Sort-Object priority
    [PSCustomObject]$ConditionEAPHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionEAP -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}
else {
    $ConditionNOEAP="NO matching Group Email Address Policy for the groups on the organization"
    [PSCustomObject]$ConditionEAPHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionNOEAP
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}
 }
else {
    [PSCustomObject]$ConditionEAPHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataString $ConditionNOEAP
    $null = $TheObjectToConvertToHTML.Add($ConditionEAPHTML)
}

#endregion Check if Distribution Group can't be upgraded because EmailAddressPolicyViolated

#region Check if Distribution Group can't be upgraded because DlHasParentGroups
[string]$SectionTitle = "Validating Distribution Group Parent Membership"
[string]$Description = "Checking if Distribution Group can't be upgraded because it is a child group of another parent group"
$ConditionParentDG=@()
try {
    Write-Host "Retrieving all distribution groups in Exchange online, please wait...." -ForegroundColor Yellow
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
$DGcounter=0
foreach($parentdg in $alldgs)
{
    try {
        $DistributionGroupMembers = Get-DistributionGroupMember $($parentdg.Guid.ToString()) -ErrorAction Stop -ResultSize unlimited
        if($alldgs.count -ge 2)
        {
            $DGcounter++
            $percent=[Int32]($DGcounter/$alldgs.count*100)
            Write-Progress -Activity "Querying Distribution Groups"  -PercentComplete $percent -Status "Processing $DGcounter/$($alldgs.count)group"
        }
    }
    catch {
        $CurrentProperty = "Retrieving: $parentdg members"
        $CurrentDescription = "Failure"
        write-log -Function "Retrieve Distribution Group membership" -Step $CurrentProperty -Description $CurrentDescription
    }
    $DGmembercounter=0
    foreach($DistributionGroupMember in $DistributionGroupMembers)
    {
        if ($DistributionGroupMember.Alias -eq $dg.Alias)
        {
        $parentdgcount++
        $ConditionParentDG=$ConditionParentDG+$parentdg
        }
        if ($DistributionGroupMembers.count -ge 2) {
        $DGmembercounter++
        $childpercent=[Int32]($DGmembercounter/$DistributionGroupMembers.count*100)
        Write-Progress -Activity "Querying Group Members" -Id 1 -PercentComplete $childpercent -Status "Processing $DGmembercounter/$($DistributionGroupMembers.count) member"    
        } 
    }
}
Write-Progress -Activity "Querying Group Members" -Completed -Id 1
Write-Progress -Activity "Querying Distribution Groups" -Completed
if($parentdgcount -le 1)
{
    [String]$NoParentDG="Distribution group is NOT a member of another group"
    [PSCustomObject]$ConditionParentDGHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $NoParentDG
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
else {
    $ConditionParentDG = $ConditionParentDG |Select-Object @{ Name = 'Parent Group Display Name';  Expression = {$_.DisplayName}},@{ Name = 'Parent Group Alias';  Expression = {$_.Alias}},@{ Name = 'Parent Group GUID';  Expression = {$_.GUID}},@{ Name = 'Parent Group RecipientTypeDetails';  Expression = {$_.RecipientTypeDetails}},@{ Name = 'Parent Group PrimarySmtpAddress';  Expression = {$_.PrimarySmtpAddress}}
    [PSCustomObject]$ConditionParentDGHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionParentDG -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionParentDGHTML)
}
#endregion Check if Distribution Group can't be upgraded because DlHasParentGroups

#region Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes with RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser
[string]$SectionTitle = "Validating Distribution Group Members Recipient Types"
[string]$Description = "Checking if Distribution Group can't be upgraded because DL contains member RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser"
try {
    Write-Host "Retrieving $($dg.PrimarySmtpAddress) members, please wait...." -ForegroundColor Yellow
    $members = Get-DistributionGroupMember $($dg.Guid.ToString()) -ErrorAction stop -ResultSize unlimited
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
    [PSCustomObject]$ConditionDGmembersHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $matchingMbr -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)

    } 
else {
    [String]$ConditionDGmembers="Distribution group contains supported members"
    [PSCustomObject]$ConditionDGmembersHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ConditionDGmembers
    $null = $TheObjectToConvertToHTML.Add($ConditionDGmembersHTML)
    }

#endregion Check if Distribution Group can't be upgraded because DlHasNonSupportedMemberTypes with RecipientTypeDetails other than UserMailbox, SharedMailbox, TeamMailbox, MailUser

#region Check if Distribution Group can't be upgraded because it has more than 100 owners or it has no owner
[string]$SectionTitle = "Validating Distribution Group Owners Count"
[string]$Description = "Checking if Distribution Group can't be upgraded because it has more than 100 owners or it has no owners"
$ConditionDGowners=New-Object PSObject
$owners=$dg.ManagedBy
if ($owners.Count -gt 100) {
    #add check to enter below region in case there are owners to check their mailboxes status
    $checkifownerhasmailbox="Continuechecking"
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "Owners are greater than 100"
    [PSCustomObject]$ConditionDGownersHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
} 
elseif ($owners.Count -eq 0) {
    $ConditionDGowners|Add-Member -NotePropertyName "Owners Count" -NotePropertyValue "No owners found"
    [PSCustomObject]$ConditionDGownersHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGowners -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
else {
    #add check to enter below region in case there are owners to check their mailboxes status
    $checkifownerhasmailbox="Continuechecking"
    $DGownersfound="Distrubtion group Owners found and are less than 100"
    [PSCustomObject]$ConditionDGownersHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $DGownersfound
    $null = $TheObjectToConvertToHTML.Add($ConditionDGownersHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution list which has more than 100 owners or it has no owner

#region Check if Distribution Group can't be upgraded because the distribution list owner(s) is non-supported with RecipientTypeDetails other than UserMailbox, MailUser
if ($checkifownerhasmailbox -match "Continuechecking")
{
    [string]$SectionTitle = "Validating Distribution Group Owners RecipientTypeDetails"
    [string]$Description = "Checking if Distribution Group can't be upgraded because DL owner(s) is non-supported with RecipientTypeDetails other than UserMailbox, MailUser"
    $ConditionDGownernonsupported=@()
    $ConditionDGownernonsupportedforusers=@()
    foreach($owner in $owners)
    {
        try {
            $owner=Get-Recipient $owner -ErrorAction stop -ResultSize unlimited
            if (!($owner.RecipientTypeDetails -eq "UserMailbox" -or $owner.RecipientTypeDetails -eq "MailUser")) 
                { 
                    $ConditionDGownernonsupported=$ConditionDGownernonsupported+$owner
                }
            }
        catch {
            $CurrentProperty = "Validating: $owner RecipientTypeDetails"
            $CurrentDescription = "Failure"
            write-log -Function "Validate owner RecipientTypeDetails" -Step $CurrentProperty -Description $CurrentDescription
            #check if the owner RecipientTypeDetails is User
            $owner=Get-User $owner -ErrorAction stop -ResultSize unlimited
            $ConditionDGownernonsupportedforusers=$ConditionDGownernonsupportedforusers+$owner
        }
    }

    if($ConditionDGownernonsupported.Count -ge 1 -or $ConditionDGownernonsupportedforusers.Count -ge 1)
    {
        $ConditionDGownernonsupported=$ConditionDGownernonsupported+$ConditionDGownernonsupportedforusers
        $ConditionDGownernonsupported=$ConditionDGownernonsupported|Select-Object Name,GUID,RecipientTypeDetails
        [PSCustomObject]$ConditionDGownernonsupportedHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGownernonsupported -TableType "Table"
        $null = $TheObjectToConvertToHTML.Add($ConditionDGownernonsupportedHTML)
    }
    else {
        $ownershaveMBXs="Distrubtion group Owner(s) RecipientTypeDetails is supported"
        [PSCustomObject]$ConditionDGownernonsupportedHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $ownershaveMBXs
        $null = $TheObjectToConvertToHTML.Add($ConditionDGownernonsupportedHTML)
    }

}
else {
    #There are no owners found
}
#endregion Check if Distribution Group can't be upgraded because the distribution list owner(s) is non-supported with RecipientTypeDetails other than UserMailbox, MailUser

#region Check if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL
[string]$SectionTitle = "Validating Distribution Group Sender Restriction"
[string]$Description = "Checking if Distribution Group can't be upgraded because the distribution list is part of Sender Restriction in another DL"
$ConditionDGSender=@()
$DGcounterloop=0
[int]$SenderRestrictionCount=1
foreach($alldg in $alldgs)
{
    if($alldgs.count -ge 2)
        {
            $DGcounterloop++
            $perc=[Int32]($DGcounterloop/$alldgs.count*100)
            Write-Progress -Activity "Validating Distribution Groups Sender Restriction"  -PercentComplete $perc -Status "Processing $DGcounterloop/$($alldgs.count)group"
        }
if ($alldg.AcceptMessagesOnlyFromSendersOrMembers -like $dg.Name -or $alldg.AcceptMessagesOnlyFromDLMembers -like $dg.Name )
{
    
    $ConditionDGSender=$ConditionDGSender+$alldg
    $SenderRestrictionCount++
}
}
Write-Progress -Activity "Validating Distribution Groups Sender Restriction" -Completed
if ($SenderRestrictionCount -le 1) {
    $NoDGSenderfound="Distribution group is NOT part of Sender Restriction in another group"
    [PSCustomObject]$ConditionDGSenderHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $NoDGSenderfound
    $null = $TheObjectToConvertToHTML.Add($ConditionDGSenderHTML)
}
else {
    $ConditionDGSender=$ConditionDGSender|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
    [PSCustomObject]$ConditionDGSenderHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ConditionDGSender -TableType "Table"
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
    [PSCustomObject]$ConditionnonsupportedrecHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditionnonsupportedrec -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
else {
    $supportedrec="Distribution group isn't a Security group nor a Dynamic distribution group nor converted to a RoomList"
    [PSCustomObject]$ConditionnonsupportedrecHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $supportedrec
    $null = $TheObjectToConvertToHTML.Add($ConditionnonsupportedrecHTML)
}
#endregion Check if Distribution Group can't be upgraded because Distribution lists which were converted to RoomLists or isn't a security group nor Dynamic DG

#region Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox
$Conditionfwdmbx=@()
[string]$SectionTitle = "Validating Distribution Group Forwarding Usage"
[string]$Description = "Checking if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox"

try {
    Write-Host "Retrieving all shared mailboxes in Exchange online, please wait...." -ForegroundColor Yellow
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
$Sharedcounter=0
foreach($sharedMBX in $sharedMBXs)
{
    if ($sharedMBX.ForwardingAddress -eq $dg.name -or $sharedMBX.ForwardingSmtpAddress -eq $dg.PrimarySmtpAddress)
    {
        $Conditionfwdmbx= $Conditionfwdmbx+$sharedMBX
        $counter++
    }
    if ($sharedMBXs.count -ge 2) 
    {
        $Sharedcounter++
        $percent=[Int32]($Sharedcounter/$sharedMBXs.count*100)
        Write-Progress -Activity "Querying Shared Mailboxes"  -PercentComplete $percent -Status "Processing $Sharedcounter/$($sharedMBXs.count) Mailboxes"
    }
    
}
Write-Progress -Activity "Querying Shared Mailboxes" -Completed
if ($counter -le 1) {
    $Nofwdmbxfound="Distribution group is NOT configured to be a forwarding address for any Shared Mailbox"
    [PSCustomObject]$ConditionfwdmbxHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $Nofwdmbxfound
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
else {
    $Conditionfwdmbx=$Conditionfwdmbx|Select-Object DisplayName,Alias,GUID,RecipientTypeDetails,PrimarySmtpAddress
    [PSCustomObject]$ConditionfwdmbxHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditionfwdmbx -TableType "Table"
    $null = $TheObjectToConvertToHTML.Add($ConditionfwdmbxHTML)
}
#endregion Check if Distribution Group can't be upgraded because the distribution list is configured to be a forwarding address for Shared Mailbox

#region Check for duplicate Alias,PrimarySmtpAddress,Name,DisplayName on EXO objects
$Conditiondupobj=@()
[string]$SectionTitle = "Validating Distribution Group Duplicates"
[string]$Description = "Checking if Distribution Group can't be upgraded because duplicate objects having same Alias,PrimarySmtpAddress,Name,DisplayName found"
try {
    Write-Host "Querying across Exchange online recipients for duplicate objects with $($dg.PrimarySmtpAddress) group, please wait..." -ForegroundColor Yellow
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
        [PSCustomObject]$ConditiondupobjHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Conditiondupobj -TableType "Table"
        $null = $TheObjectToConvertToHTML.Add($ConditiondupobjHTML)
    }
    else {
        $Nodupobjfound="No Duplicate objects found sharing same Alias,PrimarySmtpAddress,Name & DisplayName"
        [PSCustomObject]$ConditiondupobjHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDataString $Nodupobjfound
        $null = $TheObjectToConvertToHTML.Add($ConditiondupobjHTML)    
    }


#endregion Check for duplicate Alias,PrimarySmtpAddress,Name,DisplayName on EXO objects


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
