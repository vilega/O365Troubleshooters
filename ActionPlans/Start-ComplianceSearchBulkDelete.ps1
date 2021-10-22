<#
        .SYNOPSIS
        Delete more than 10 items using Compliance Search Action - Single mailbox scenario 

        .DESCRIPTION
        Provide an automated way to delete more than 10 items from a particular mailbox using Compliance Search Action.
        This works only for Compliance Searches targeting a single mailbox.
        
        .EXAMPLE
        Create a Compliance Search from Compliance Portal to target a specific mailbox, inspect the search results to confirm your search criteria returned the expected items that you wish to delete.
        If the returned items count is more than 10, you can start this script to automate the deletion of those items.
        
        .LINK
        Online documentation: https://aka.ms/O365Troubleshooters/ComplianceSearchActionBulkDelete

    #>

# Function to check&wait until Compliance Search finishes running
function Test-SearchStatusIsComplete {
    param ($SearchNameToTest)
    Write-Host "> Checking current status for the selected Compliance Search '$searchname' and waiting to complete in case it is still running."
    Do {
        Write-Host '*' -NoNewline
        Start-Sleep -Seconds 5
        $search = Get-ccComplianceSearch $SearchNameToTest
    } Until (($search.Status -eq 'Completed') -and ($search.JobProgress -eq '100'))
    Start-Sleep -Seconds 5
    Write-Host "`n> Compliance Search status is $($search.Status) and its progress is $($search.JobProgress)%"
}

$cleanup = {
    Write-Host "Currently the search is configured to exclude the 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders from its scope."
    Write-Host -ForegroundColor Cyan "How do you prefer to keep the search?"
    Write-Host "Type 'K' to KEEP the exclusion: If you re-run the search, it will NOT find the initial items, but they are available under 'Recoverable Items' mailbox folder, governed by the mailbox SIR settings and holds applied."
    Write-Host
    Write-Host "Type 'R' to Remove the exclusion: If you re-run the search, it will find the initial items, but they are available under 'Recoverable Items' mailbox folder, governed by the mailbox SIR settings and holds applied."
    
    Do {
        [string]$option = read-host "Please type 'K' or 'R' to select one of the above mentioned options"
        $option = $option.ToLower()
    } Until (($option -eq "k") -or ($option -eq "r"))
        
    If ($Option -eq "r") { 
        Write-Host -ForegroundColor Yellow "Reverting to initial search folder scope for Compliance search $searchname"
        Set-ccComplianceSearch $searchname -ContentMatchQuery $OldContentMatchQuery
        Test-SearchStatusIsComplete $searchname
        # Add to execution report
        [string]$SectionTitle = "Updated Compliance Search Query - remove exclusions"
        [string]$Description = "Selected '$searchname' search query has been updated to remove the exclusions and revert to the initial search query: '$OldContentMatchQuery'"
        [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
        $null = $TheObjectToConvertToHTML.Add($SectionHTML)
        Write-Log -function "Start-complianceSearchBulkDelete" -step  "Opting to Keep exclusion for recoverable items folders or Revert to initial search query" -Description "Selected '$searchname' search query has been updated to remove the exclusions and revert to the initial search query: '$OldContentMatchQuery'"
    }
    Else {
        Write-Host -ForegroundColor Yellow "Compliance search $searchname remains configured to exclude the 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders from its scope."
        $ComplianceSearch = Get-ccComplianceSearch -Identity $searchname
        $latestquery = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
        # Add to execution report
        [string]$SectionTitle = "Updated Compliance Search Query - keep exclusions"
        [string]$Description = "Selected '$searchname' search query has been updated to keep the exclusions: '$latestquery'"
        [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
        $null = $TheObjectToConvertToHTML.Add($SectionHTML)
        Write-Log -function "Start-complianceSearchBulkDelete" -step  "Opting to Keep exclusion for recoverable items folders or Revert to initial search query" -Description "Selected '$searchname' search query is keeping the exclusions and the current search query is: '$latestquery'"
    }
}  

Clear-Host

# Connect to required O365 services via PowerShell
$Workloads = "exo", "SCC"
Connect-O365PS $Workloads

# Initiate object for execution report
$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"

# Display warnings !!!
Clear-Host
Write-Host -ForegroundColor Red "IMPORTANT:"
Write-Host -ForegroundColor Red -BackgroundColor White "This tool is intended for single mailbox Compliance Searches only!"
Write-Host -ForegroundColor Red -BackgroundColor White "Trying to adjust it for multiple mailboxes may lead to unforeseen issues, due to per tenant Compliance Search limits!"
Write-Host -ForegroundColor Red "IMPORTANT:"
Write-Host -ForegroundColor Yellow "If there are no holds protecting the items or mailbox, there is the risk for these items to be purged with the next Managed Folder Assistant run!"
Write-Host "Managed Folder Assistant runs automatically at anytime between 1 to 7 days since its last execution."
Write-Host
Write-Host "If there are holds protecting the items or the mailbox, the items will be present under 'Recoverable Items' mailbox folder after deletion process finishes. They will not be accessible to the user via email clients, but the tenant admin will be able to either restore them or find them using Compliance Search and export them as PST."   

# Select the search from existing searches

$allSearches = Get-ccComplianceSearch
Write-Host
Write-Host -ForegroundColor Cyan "Please select the Compliance Search scoped for a single mailbox for which you wish to delete the found items:"
[string]$searchname = ($allSearches | Select-Object name | Out-GridView -OutputMode single -Title "Select one search").Name

# Wait for Compliance Search in case it is running
Test-SearchStatusIsComplete $searchname

# Validate if search 'locations' contains only one mailbox, else prompt 'Not single mailbox search, please select a search for a single mailbox' and loop at beginning.
$ComplianceSearch = Get-ccComplianceSearch -Identity $searchname
If ($compliancesearch.ExchangeLocation.count -ne 1) {
    Write-Host
    write-host "You have selected a Compliance Search scoped for more than 1 mailbox, please press any key to restart and select a search scoped for a single mailbox."
    Write-Log -function "Start-complianceSearchBulkDelete" -step  "Selecting a Compliance Search scoped for a single mailbox" -Description "Selected the Compliance Search '$ComplianceSearch', which is scoped for more than 1 mailbox, redirecting to new search selection."
    Read-Key
    Start-O365TroubleshootersMenu
}

# Get details for selected Compliance Search
$initialitems = $ComplianceSearch.Items
$query = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
$contentsize = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].contentsize
[int64]$contentsizebytes = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].contentsizeRaw
$location = $compliancesearch.ExchangeLocation

# Get Single Item Recovery (SIR) settings present on mailbox
[bool]$SIREnabled = (Get-Mailbox $location[0]).SingleItemRecoveryEnabled
$SIRRetainDays = (Get-Mailbox $location[0]).RetainDeletedItemsFor.split(".")[0]

# Get Recoverable Items relevant mailbox values
[int64]$RecoverableItemsSize = (Get-MailboxFolderStatistics $location[0] -FolderScope RecoverableItems)[0].FolderAndSubfolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")
[int64]$RecoverableItemsQuota = (Get-Mailbox $location[0]).RecoverableItemsQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")
[int64]$RecoverableItemsSizeFinal = $RecoverableItemsSize + $contentsizebytes
[int64]$MaxAvailable = $RecoverableItemsQuota - $RecoverableItemsSizeFinal

# Create object for reporting
$searchdetails = [PSCustomObject]@{
    Name                            = $searchname
    Items                           = $initialitems
    "Search Query"                  = $query
    "Items size"                    = $contentsize
    Mailbox                         = $location
    "Recoverable items size before" = "$RecoverableItemsSize bytes"
    "Recoverable items size after"  = "$RecoverableItemsSizeFinal bytes"
    "Recoverable items quota"       = "$RecoverableItemsQuota bytes"
}

# Add Compliance Search details to execution report
[string]$SectionTitle = "Selected Compliance Search details"
[string]$Description = "Summary of the selected Compliance Search '$searchname'"
[PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $searchdetails -tabletype List 
$null = $TheObjectToConvertToHTML.Add($SectionHTML)

# Export Search details to csv file
$ts = get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\ComplianceSearch_bulkdelete_$ts"
mkdir $ExportPath -Force | out-null
$searchdetails | export-csv -Path $ExportPath\Selected_search_details.csv -NoTypeInformation 
Write-Log -function "Start-complianceSearchBulkDelete" -step  "Creating report object and exporting details for Compliance Search '$searchname' as csv file" -Description "Exported details for Compliance Search '$searchname' as csv file at path: $ExportPath\Selected_search_details.csv"

Write-Log -function "Start-complianceSearchBulkDelete" -step  "User to select the Compliance Search" -Description "User selected the Compliance Search '$searchname'"

Write-Host -ForegroundColor Yellow "You have selected the Compliance Search named '$Searchname', having $initialitems items found with their total size of $contentsize in mailbox '$location'."

# Check for Single Item Recovery (SIR) settings on mailbox

If ($SIRRetainDays.Substring(0, 1) -eq "0") {
    Write-Host -ForegroundColor DarkRed -BackgroundColor White "'RetainDeleteditemsFor' property is set to 0 days for this mailbox!!"
    Write-Host -ForegroundColor Red -BackgroundColor White "If you choose to proceed with deletion, the items will be PERMANENTLY LOST!!!"
    # Add SIR status to execution report
    [string]$SectionTitle = "Single Items Recovery (SIR) status on target mailbox"
    [string]$Description = "Single Items Recovery status on target mailbox $location"
    [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "'RetainDeleteditemsFor' is set to 0 days for mailbox '$location', the items will be PERMANENTLY LOST!!!"
    $null = $TheObjectToConvertToHTML.Add($SectionHTML)
}
Elseif ( !($SIREnabled) ) {
    Write-Host -ForegroundColor DarkRed -BackgroundColor White "'Single Item Recovery' is DISABLED for this mailbox!!"
    Write-Host -ForegroundColor Red -BackgroundColor White "If you choose to proceed with deletion, the items will be PERMANENTLY LOST!!!"
    # Add SIR status to execution report
    [string]$SectionTitle = "Single Items Recovery (SIR) status on target mailbox"
    [string]$Description = "Single Items Recovery status on target mailbox $location"
    [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "'Single Item Recovery' is DISABLED for mailbox '$location', the items will be PERMANENTLY LOST!!!"
    $null = $TheObjectToConvertToHTML.Add($SectionHTML)
}
Else {
    Write-Host -ForegroundColor Yellow "'RetainDeletedItemsFor' parameter is set to $SIRRetainDays days for this mailbox!"
    Write-Host -ForegroundColor Yellow "After deletion finishes, the items will be PERMANENTLY LOST after $SIRRetainDays days!!"
    # Add SIR status to execution report
    [string]$SectionTitle = "Single Items Recovery (SIR) status on target mailbox"
    [string]$Description = "Single Items Recovery status on target mailbox $location"
    [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "'RetainDeletedItemsFor' parameter is set to $SIRRetainDays days for mailbox '$location', the items will be PERMANENTLY LOST after $SIRRetainDays days!!"
    $null = $TheObjectToConvertToHTML.Add($SectionHTML)
}
 
# Check available space under Recoverable Items

If (($MaxAvailable -le 0) -and ($SIREnabled -and ($SIRRetainDays.Substring(0, 1) -ne "0"))) {
    Write-Host "Not enough space available under 'Recoverable Items' folder to accommodate the items!"
    Write-Host "Either adjust the search query to return fewer items, or make additional space in 'Recoverable Items' mailbox folder."
    Read-Key
    Start-O365TroubleshootersMenu
}
Elseif ($SIREnabled -and ($SIRRetainDays.Substring(0, 1) -ne "0")) {
    Write-Host "At the end of deleting the $initialitems items, mailbox '$location' will have $MaxAvailable bytes of free space remaining available under 'Recoverable Items' folder."
    If ($MaxAvailable -le 1073741824) {
        Write-Host -ForegroundColor Yellow "The remaining space under 'Recoverable Items' folder will be below 1GB after deletion!"
        Write-Host -ForegroundColor Red "IMPORTANT:"
        Write-Host -ForegroundColor Yellow "Please avoid reching quota for 'Recoverable Items', to avoid issues such as user not being able to delete emails!"
        # Add Recoverable Items warning to execution report
        [string]$SectionTitle = "Low space remaining for 'Recoverable Items' folder!"
        [string]$Description = "Low space remaining for 'Recoverable Items' folder!"
        [PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "Less than 1GB space left in 'Recoverable Items' folder for mailbox $location! Please avoid reching quota for 'Recoverable Items', to avoid issues such as user not being able to delete emails!"
        $null = $TheObjectToConvertToHTML.Add($SectionHTML)
    }
}

# Confirmation input to proceed with deletion
Write-Host -ForegroundColor Cyan "Are you sure you want to delete the $initialitems items with size $contentsize found by the selected '$searchname' Compliance Search from mailbox '$location'?"
 
[string]$Option = read-host "Type 'yes' to confirm"
$option = $option.ToLower()
If ($Option -ne "yes") { 
    Write-Host "You haven't confirmed by typing 'yes'. Press any key to restart."
    Read-Key
    Start-O365TroubleshootersMenu 
}

# Add record of user confirmation to execution report
[string]$SectionTitle = "User confirmation choice"
[string]$Description = "User '$global:userPrincipalName' confirmed to proceed with bulk deletion of $initialitems items with size $contentsize found by the selected '$searchname' Compliance Search from mailbox '$location', by typing 'yes' +ENTER"
[PSCustomObject]$SectionHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
$null = $TheObjectToConvertToHTML.Add($SectionHTML)

Write-Log -function "Start-complianceSearchBulkDelete" -step  "User confirmation to go ahead with items deletion for Compliance Search '$searchname'" -Description "User '$global:userPrincipalName' confirmed to go ahead with items deletion for Compliance Search '$searchname'"

# Identifying 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders - this part is taken from article:  
Write-Host
Write-Host "Identifying 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders"
    
[string]$mbx = $compliancesearch.exchangelocation 
$folderQueries = @()
$folderStatistics = Get-MailboxFolderStatistics $mbx | where-object { ($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions") -or ($_.FolderPath -eq "/DiscoveryHolds") }
foreach ($folderStatistic in $folderStatistics) {
    $folderId = $folderStatistic.FolderId;
    $folderPath = $folderStatistic.FolderPath;
    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF");
    $folderIdBytes = [Convert]::FromBase64String($folderId);
    $indexIdBytes = New-Object byte[] 48;
    $indexIdIdx = 0;
    $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
    $folderStat = New-Object PSObject
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
    $folderQueries += $folderStat
}
       
$RecoverableItemsFolder = $folderQueries.folderquery[0]
$PurgesFolder = $folderQueries.folderquery[1]
$VersionsFolder = $folderQueries.folderquery[2]
$DiscoveryHoldsFolder = $folderQueries.folderquery[3]

# Adjusting the search scope to exclude 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders 
Write-Host -ForegroundColor Yellow "Adjusting the search scope to exclude 'Recoverable Items', 'Purges', 'DiscoveryHolds' and 'Versions' folders"
   
[string]$OldContentMatchQuery = $ComplianceSearch.ContentMatchQuery
[string]$NewContentMatchQuery = $null
[string]$NewContentMatchQuery = [string]$OldContentMatchQuery + "(NOT (($RecoverableItemsFolder) OR ($PurgesFolder) OR ($VersionsFolder)  OR ($DiscoveryHoldsFolder)))"
    
Set-ccComplianceSearch $searchname -ContentMatchQuery $NewContentMatchQuery

# Wait for Compliance Search in case it is running
Test-SearchStatusIsComplete $searchname

# Calculate the number of batches
$Iterations = [math]::Ceiling($initialitems / 10)
$iteration = 0

Write-Log -function "Start-complianceSearchBulkDelete" -step  "Identifying and excluding 'Recoverable Items'and 'Purges', 'DiscoveryHolds' and 'Versions' folders from Compliance Search '$searchname'" -Description "Identifyied and excluded 'Recoverable Items'and 'Purges', 'DiscoveryHolds' and 'Versions' folders from Compliance Search '$searchname'"

# Starting bulk deletion
    
Write-Host
write-host "Deleting " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems" -NoNewline; Write-Host -ForegroundColor white " items in " -NoNewline; Write-Host -ForegroundColor yellow "$iterations" -NoNewline; Write-Host -ForegroundColor white " batches of 10 items each, due to Compliance Search Action limit"
    
DO {
    $iteration++

    # Check for and delete any existing purge actions for the selected Compliance Search
    $PurgeAction = Get-ccComplianceSearchAction | Where-Object { ($_.Name -match "$searchname") -and ($_.Name -match "_Purge") }
    If ($PurgeAction) {
        foreach ($p in $PurgeAction) { Remove-ccComplianceSearchAction ($p).name -Confirm:$False | Out-Null }
    }
                            
    Start-Sleep -Seconds 5
        
    # Start the batches
    Write-Host -ForegroundColor Yellow "Batch no. [$Iteration / $Iterations]"
        
    # Create Purge Compliance Search Action for the selected Compliance Search
    [boolean]$repeat = $True
    $i = 1
    while ($repeat) {
            
        try {
            New-ccComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete -Confirm:$false -ErrorAction stop | Out-Null
            $repeat = $false
        }
        catch {
            Start-Sleep -Seconds 5
            Write-Log -function "Start-ComplianceSearchBulkDelete" -step  "Create new Compliance Search Action" -Description "$($PSItem.Exception.Message)"
            if ($i -lt 6) {
                $i++
            } 
            else {
                Write-Host "cannot create new search action, returning to menu"
                Read-Key
                &$cleanup
                Start-O365TroubleshootersMenu
            }
        }
    }   
        
    Write-Host "> current batch Purge Action status: Running"
    Do {
        $PurgeAction = Get-ccComplianceSearchAction -Identity "$searchname`_Purge"
        Write-Host '*' -NoNewline
        Start-Sleep -Seconds 5
    } While ($PurgeAction.Status -ne 'Completed')
    Write-Host "`n> current batch Purge Action status: Completed"
    
} While ($iteration -ne $Iterations)

Write-Host
Write-Host -ForegroundColor Yellow "Finished deleting $initialitems items with size $contentsize found by the selected $searchname Compliance Search!"

# Remove Compliance Search Action remained after last batch
Remove-ccComplianceSearchAction ($PurgeAction).name -Confirm:$False | Out-Null
    
# Call Cleanup script block to offer choice to keep or remove recoverable folders exclusions from the Content Search
&$cleanup
    
#Build HTML report out of the previous HTML sections
    
[string]$FilePath = $ExportPath + "\Selected_search_report.html"
    
Export-ReportToHTML -FilePath $FilePath -PageTitle "Compliance Search Bulk Delete Action execution report" -ReportTitle "Compliance Search Bulk Delete Action execution report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
    
#Ask end-user for opening the HTMl report
    
$OpenHTMLfile = Read-Host "Do you wish to open HTML report file now?`nType 'Y' to open or any other character to exit!"
    
if ($OpenHTMLfile.ToLower() -like "*y*") {
    Write-Host "Opening report...." -ForegroundColor Cyan
    Start-Process $FilePath
}
    
#endregion ResultReport
    
# Print location where the data was exported
    
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
    
Read-Key 

Write-Log -function "Start-complianceSearchBulkDelete" -step  "Return to main menu" -Description "Done"

Start-O365TroubleshootersMenu
    
    