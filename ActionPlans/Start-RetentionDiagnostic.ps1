Function Confirm-User {
    $global:path = "C:\Users\vilega\OneDrive - Microsoft\EEE\Torus\Temp\$bug"
    mkdir $path -Force | out-null
    Set-Location $path

    #$SaveErrorActionPreference = $ErrorActionPreference
    #$ErrorActionPreference = SilentlyContinue
   
    $global:User = get-user $u -ErrorAction SilentlyContinue
    if ($User) {
        
        Write-Host "$u is a $($User.RecipientTypeDetails)"
        if ($User.RecipientType -eq "UserMailbox") {   
            $Global:m = Get-mailbox $u

        } 
        elseif ($User.RecipientType -eq "MailUser") {
            $Global:m = Get-mailuser $u
        }


        $Global:MailboxLocations = Get-MailboxLocation -User $($m.Guid.Guid)

        $Global:MailboxLocations =  $MailboxLocations | ?{($_.MailboxLocationType -eq "Primary") -or ( $_.MailboxLocationType -eq "MainArchive") -or ( $_.MailboxLocationType -eq "AuxArchive")}
        Write-Host "User $u has the $($MailboxLocations.count) mailbox locations"
        $Global:PrimaryMailbox = $MailboxLocations | ? MailboxLocationType -eq "Primary"
        $Global:MainArchive = $MailboxLocations | ? MailboxLocationType -eq "MainArchive"



    }
    else {

        Write-Host "User cannot be found!" -ForegroundColor Red
        Read-Host "Press [ENTER] to reload the main menu"
        Show-Menu
    }
}

Function Check-DemotedArchive{
    foreach ($MailboxLocation in $MailboxLocations) {
        if ($MailboxLocation.MailboxLocationType -eq "DemotedArchive") {
            Write-Host "User $u has a demoted archive $($MailboxLocation.Identify)" -ForegroundColor Red
            # Customer can remove them if needed by using Set-Mailbox -RemoveDisabledArchive / set-mailuser -RemoveDisabledArchive
        }
    }
}

Function Check-MailboxStatistics {

    [System.Collections.ArrayList]$global:MailboxStatistics = @()
    foreach ($MailboxLocation in $m.MailboxLocations) {
        $CurrentMailboxStats = get-mailboxstatistics "$($MailboxLocation.split(";")[1])"
        $null = $MailboxStatistics.Add($CurrentMailboxStats)
    }
    $MailboxStatistics | ft MailboxGuid, Total*Size, MailboxTypeDetail, IsArchiveMailbox
    
    if ($m.MailboxLocations -match "MainArchive") {
        Write-Host "Total Archive Mailbox Statistics" -ForegroundColor Cyan
        $Global:ArchiveMailboxStatistics = Get-MailboxStatistics $u -archive
        $ArchiveMailboxStatistics | fl MailboxGuid, Total*Size, MailboxTypeDetail, IsArchiveMailbox, DeletedItemCount, ItemCount
    }
}

Check-MRMforUser{
    If (($OrganizationConfig).ElcProcessingDisabled -eq $true) {
        Write-Host "ELC is disabled at Tenant level!" -ForegroundColor Red
    }
    RetentionHoldEnabled

    Write-Host "`nUser license:" -ForegroundColor Cyan
    $m.PersistedCapabilities # Exchange_Enterprise, Archive-Addon

    Write-Host "`nIf account is disabled, EWS is disable or Self permissions are missing the items are not moved to archive. Self needs both {FullAccess, ReadPermission}" -ForegroundColor Cyan
    
    # AccountDisabled 
    $m | fl AccountDisabled, ExchangeUserAccountControl, RecipientType*


    # EWS Check
    get-casmailbox $o\$u | fl *ews*

    # Mailbox Permissions
    Get-MailboxPermission $o\$u -User Self

    Write-Host "`nIf there are items larger than max send/receive size, these won't be archived and can block the archive split if there are already in MainArchive" -ForegroundColor Cyan
    $m | fl max*size

    Write-Host "`nMailbox quotas:" -ForegroundColor Cyan
    $mb | fl *quota*, UseDatabaseQuotaDefaults

    Write-Host "`nCheck if ELC is proccessing the mailbox, items are moved to Purges and when expire" -ForegroundColor Cyan
    $m | fl ElcProcessingDisabled, RetentionHoldEnabled , SingleItemRecoveryEnabled, RetainDeletedItemsFor, UseDatabaseRetentionDefaults, UseDatabaseQuotaDefaults

}
Function Check-Holds {
    Write-Host "`nChecking Organization's Holds!" -ForegroundColor Cyan
    $OrganizationConfig = get-organizationconfig $o
    If (($OrganizationConfig).InPlaceHolds) {
        Write-Host "Organization has the following global holds: " -ForegroundColor Yellow
        ($OrganizationConfig).InPlaceHolds
    }
    else {
        Write-Host "Organization has no global holds: " -ForegroundColor Yellow
    }



    Write-Host "`nChecking Holds settings for user $u !" -ForegroundColor Cyan
    $UserHolds = $m | select LitigationHoldEnabled, EndDateForRetentionHold, StartDateForRetentionHold, LitigationHoldDate, LitigationHoldOwner, ComplianceTagHoldApplied, DelayHoldApplied, DelayReleaseHoldApplied, LitigationHoldDuration, InPlaceHolds

}

Function Check-LastErrorComponent {

    Write-Host "`nLast MRM error for mailbox: $u" -ForegroundColor Cyan
    $MRMMailboxLog = Export-MailboxDiagnosticLogs $o\$u -ComponentName MRM
    $MRMMailboxLog 
} 

Function Check-LegacyMRM {
    If (($m.RetentionHoldEnabled -eq $true) -or ($m.ElcProcessingDisabled -eq $true) -or ($OrganizationConfig.ElcProcessingDisabled -eq $true)) {
        Write-Host "`nELC won't proccess mailbox $u :" -ForegroundColor Red
        Write-Host "RetentionHoldEnabled:"
        $m.RetentionHoldEnabled
        Write-Host "ElcProcessingDisabled on $u :"
        $m.ElcProcessingDisabled 
        Write-Host "ElcProcessingDisabled at Organization Level:"
        $OrganizationConfig.ElcProcessingDisabled 
    }
    
    Write-Host "`nRetention policy: $($m.RetentionPolicy)" -ForegroundColor Cyan
    $RetentionPolicy = $m.RetentionPolicy


    If ($m.RetentionPolicy -eq $null) {
        Write-Host "`nNo RetentionPolicy assigned on the mailbox $u, this can prevent MRM config to be recreated" -ForegroundColor Red
    }
    $RetentionPolicy = Get-RetentionPolicy $m.RetentionPolicy 
    $RetentionPolicyTagLinks = $RetentionPolicy.RetentionPolicyTagLinks
    $RetentionPolicyTags = $RetentionPolicyTagLinks | Get-RetentionPolicyTag 
    Write-Host "`nThe followings RetentionPolicyTag are included in $($m.RetentionPolicy)" -ForegroundColor Cyan
    $RetentionPolicyTags | ft Name, Type, RetentionAction, AgeLimitForRetention, Guid, RetentionId


    # MRMConfiguration - MFCMapi Article
}

Function Check-FolderStatistics {
    Write-Host "Checking Folder Statistics:" -ForegroundColor Cyan
    $SelectShards = , ($MailboxLocations | Out-GridView -Title "Select the Shards from which you need to check MailboxFolderStatistics:" -PassThru)
    $Parameters = "IncludeOldestAndNewestItems", "IncludeRecoverableItems", "FolderScope", "IncludeAnalysis"
    $FolderScope = @(
        "All",
        "Archive",
        "Calendar",
        "Contacts",
        "ConversationHistory",
        "DeletedItems",
        "Drafts",
        "Inbox",
        "JunkEmail",
        "Journal",
        "LegacyArchiveJournals",
        "ManagedCustomFolder",
        "NonIpmRoot",
        "Notes",
        "Outbox",
        "Personal",
        "RecoverableItems",
        "RssSubscriptions",
        "SentItems",
        "SyncIssues",
        "Tasks")
    $selectParameters = , ($Parameters | Out-GridView -Title "Select the parameters to use for MailboxFolderStatistics:" -PassThru)
    If ($selectParameters -contains "FolderScope") {
        $SelectFolderScope = $FolderScope | Out-GridView -Title "Select one folder type to check Folder:" -OutputMode Single
    }

    foreach ($SelectShard in $SelectShards ) {
        $command = "Get-MailboxFolderStatistics $o\$($SelectShard.MailboxGuid.ToString())"
        foreach ($selectParameter in $selectParameters ) {
            if ($selectParameter -eq "FolderScope") {
                $command += " -FolderScope $SelectFolderScope"
            }
            else {
                $command += " -$selectParameter"
            }
        }
    

        $global:FolderStatistics = Invoke-Expression $command

        $global:FolderStatistics | % {
            $FolderSizeGB = [math]::Round(($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1gb), 2)
            $_ | Add-Member -TypeName NoteProperty -NotePropertyName FolderSizeGB -NotePropertyValue $FolderSizeGB
        }
    
        $global:FolderStatistics = $FolderStatistics | sort -descending FolderSizeGB 
    
        If ($selectParameters -contains "IncludeAnalysis") {
            $FolderStatistics | select FolderPath, FolderType, ContainerClass, ContentMailboxGuid, LastMoved*, Movable, FolderSizeGB, FolderSize, ItemsInFolder, OldestItemReceivedDate, OldestItemLastModifiedDate, ArchivePolicy, RetentionFlags, DeletePolicy, TopSubjectSize, TopSubjectCount, TopSubjectReceivedTime |
            Out-GridView -Title $command
        }
        If ($selectParameters -notcontains "IncludeAnalysis") {
            $FolderStatistics | select FolderPath, FolderType, ContainerClass, ContentMailboxGuid, LastMoved*, Movable, FolderSizeGB, FolderSize, ItemsInFolder, OldestItemReceivedDate, OldestItemLastModifiedDate, ArchivePolicy, RetentionFlags, DeletePolicy |
            Out-GridView -Title $command
        }
    }


}

Function Compliance {
    $policyName = ""
    $policy = Get-RetentionCompliancePolicy -Organization $o -Identity $policyName -distributiondetail -RetentionRuleTypes
    #$policy = Get-RetentionCompliancePolicy -Organization $o -Identity $policyName
    $policyId = $policy.ExchangeObjectId.Guid.ToString()
    $policy | fl name, ExchangeObjectId, Identity, Distribution*, Mode, RetentionRuleTypes, ExchangeLocation, ExchangeLocationException, TeamsChatLocation, TeamsChatLocationException, LastStatusUpdateTime, Enabled, LastStatusUpdateTime
    $rules = Get-RetentionComplianceRule -Organization $o -Policy $policyId
    #$rule = Get-AppRetentionComplianceRule -Organization $o -Policy $policyId 
    #Get-AppRetentionCompliancePolicy -Organization $o $policyId -DistributionDetail |fl Name, Distribution*
    $bindings = Get-ComplianceBinding -Organization $o | Where-Object { $_.policyId -eq $policyId }

    # Checking distribution errors
    $policy | select -ExpandProperty DistributionResults 
    $policy | select -ExpandProperty DistributionResults | clip
($policy | select -ExpandProperty DistributionResults).Endpoint
}

Set-GlobalVariables
Check-MailboxStatistics
Check-Holds
Check-LastErrorComponent 
Check-LegacyMRM
Check-FolderStatistics ( + PendingRescan, NeedsRescan, IncludeAnalysis on a specific folder/all)


# what other MRM configuration we can find
Get-MailboxUserConfiguration -Mailbox $o\$($MainArchive.mailboxguid.tostring())  -Identity Root\AuxArchiveFolderSplitState
$conf = Get-MailboxUserConfiguration -Mailbox $u -Identity configuration\*
$conf 

$conf = Get-MailboxUserConfiguration -Mailbox $u -Identity root\*
$conf

# ? 1M/3M folder item limits

# Ghosted Folder (MainArchive), optional on a specific shard

# quota: mailbox, send/receive, items limit in folders

# RetentionCompliance Policies;

# ? UserPolicies. Compliance check what appusercompliancepolicy applies





