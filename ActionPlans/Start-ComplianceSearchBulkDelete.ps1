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

$cleanup = {
    Write-Host "Currently the search is configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
    Write-Host -ForegroundColor Cyan "How do you prefer to keep the search?"
    Write-Host "Type 'K' to KEEP the exclusion: If you re-run the search, it will NOT find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    Write-Host
    Write-Host "Type 'R' to Remove the exclusion: If you re-run the search, it will find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    
    Do {
        [string]$option = read-host "Please type one of the above mentioned options"
        $option = $option.ToLower()
    } Until (($option -eq "k") -or ($option -eq "r"))
        
    If ($Option -eq "r") { 
        Write-Host -ForegroundColor Yellow "Reverting to initial search folder scope for Compliance search $searchname"
        Set-ccComplianceSearch $searchname -ContentMatchQuery $OldContentMatchQuery
        Do {
            $search = Get-ccComplianceSearch $searchname
            Write-Host "> current Search status is $($search.Status) and search job progress: $($search.JobProgress)%"
            Start-Sleep -Seconds 5
        } While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))
        Write-Host "Done"
    
        [string]$SectionTitle = "Updated Compliance Search Query - remove exclusions"
    
        [string]$Description = "Selected '$searchname' search query has been updated to remove the exclusions and revert to the initial search query: '$OldContentMatchQuery'"
       
        [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
          
        $null = $TheObjectToConvertToHTML.Add($SectionHTML)
       
    }
    Else {
        Write-Host -ForegroundColor Yellow "Compliance search $searchname remains configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
        $ComplianceSearch = Get-ccComplianceSearch -Identity $SelectedSearch
        $latestquery = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
    
        [string]$SectionTitle = "Updated Compliance Search Query - keep exclusions"
    
        [string]$Description = "Selected '$searchname' search query has been updated to keep the exclusions: '$latestquery'"
    
        [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
       
        $null = $TheObjectToConvertToHTML.Add($SectionHTML)
    
    }
}  
    
Clear-Host
    
$Workloads = "exo", "SCC"
       
Connect-O365PS $Workloads
    
# Select the search from existing searches
    
$allSearches = Get-ccComplianceSearch
Write-Host -ForegroundColor Yellow "Please select the Compliance Search scoped for a single mailbox for which you wish to delete the found items:"
     
[string]$SelectedSearch = ($allSearches | Select-Object name | Out-GridView -OutputMode single -Title "Select one search").Name
    
$ComplianceSearch = Get-ccComplianceSearch -Identity $SelectedSearch
    
$initialitems = $ComplianceSearch.Items
    
$searchname = $ComplianceSearch.Name
    
$query = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
    
$contentsize = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].contentsize

$location = $compliancesearch.ExchangeLocation
    
Write-Host "Found " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems " -NoNewline; Write-Host -ForegroundColor white "items for compliance Search " -NoNewline; Write-Host -ForegroundColor yellow "$searchname"
    
# Check if search 'locations' contains only one '@' symbol, else prompt 'Not single mailbox search, please select a search for a single mailbox' and loop at beginning.
If ($compliancesearch.ExchangeLocation.count -ne 1) {
    write-host "You have selected a Compliance Search scoped for more than 1 mailbox, please press any key to restart and select a search scoped for a single mailbox."
    Write-Log -function "Start-complianceSearchBulkDelete" -step  "Selecting a Compliance Search scoped for a single mailbox" -Description "Selected the Compliance Search '$ComplianceSearch', which is scoped for more than 1 mailbox, redirecting to new search selection."
    Read-Key
    Start-O365TroubleshootersMenu
}
    
# Confirmation input
    
Write-Host -ForegroundColor Cyan "Are you sure you want to delete $initialitems items with size $contentsize found by the selected $searchname Compliance Search?"
    
[string]$Option = read-host "Type 'yes' to confirm"
$option = $option.ToLower()
If ($Option -ne "yes") { 
    Write-Host "You haven't confirmed by typing 'yes'. Press any key to restart."
    Read-Key
    Start-O365TroubleshootersMenu 
}
    
# Create object for reporting
$searchdetails = [PSCustomObject]@{
    Name        = $searchname
    items       = $initialitems
    searchquery = $query
    size        = $contentsize
    mailbox     = $location
}
        
$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"
    
$ts = get-date -Format yyyyMMdd_HHmmss
    
$ExportPath = "$global:WSPath\ComplianceSearch_bulkdelete_$ts"
        
mkdir $ExportPath -Force | out-null
    
$searchdetails | export-csv -Path $ExportPath\Selected_search_details.csv -NoTypeInformation 
        
[string]$SectionTitle = "Selected Compliance Search details"
    
[string]$Description = "Summary of the selected Compliance Search '$searchname'"
    
[PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $searchdetails -tabletype List 
$null = $TheObjectToConvertToHTML.Add($SectionHTML)
    
# Identifying 'Recoverable Items', 'Purges' and 'Versions' folders - this part is taken from article:  
    
Write-Host "Identifying 'Recoverable Items', 'Purges' and 'Versions' folders"
    
[string]$mbx = $compliancesearch.exchangelocation 
$folderQueries = @()
$folderStatistics = Get-MailboxFolderStatistics $mbx | where-object { ($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions") }
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
    
# Adjusting the search scope to exclude 'Recoverable Items', 'Purges' and 'Versions' folders 
    
Write-Host -ForegroundColor Yellow "Adjusting the search scope to exclude 'Recoverable Items', 'Purges' and 'Versions' folders"
    
[string]$OldContentMatchQuery = $ComplianceSearch.ContentMatchQuery
[string]$NewContentMatchQuery = $null
[string]$NewContentMatchQuery = [string]$OldContentMatchQuery + "(NOT (($RecoverableItemsFolder) OR ($PurgesFolder) OR ($VersionsFolder)))"
    
Set-ccComplianceSearch $searchname -ContentMatchQuery $NewContentMatchQuery
Do {
    $search = Get-ccComplianceSearch $searchname
    Write-Host "> current Search status is $($search.Status)"
    Start-Sleep -Seconds 5
} While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))
        
$Iterations = [math]::Ceiling($initialitems / 10)
$iteration = 0
    
# Starting bulk deletion
    
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
    
Start-O365TroubleshootersMenu
    
    