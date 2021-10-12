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
        Online documentation: https://answers.microsoft.com/

    #>
<#
    Import-Module C:\Work\Projects\PS\GitHubStuff\O365Troubleshooters\O365Troubleshooters.psm1 -Force

# 2nd requirement Execute set global variables

Set-GlobalVariables

# 3rd requirement to start the menu

Start-O365TroubleshootersMenu

#>

$cleanup = {
    Write-Host "Currently the search is configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
    Write-Host -ForegroundColor Magenta "How do you prefer to keep the search?"
    Write-Host "Type 'K' to KEEP the exclusion. If you re-run the search, it will NOT find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    Write-Host
    Write-Host "Type 'R' to Remove the exclusion. If you re-run the search, it will find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    
    Do {
    [string]$option = read-host "Please type one of the above mentioned options"
    $option = $option.ToLower()
    } Until (($option -eq "k") -or ($option -eq "r"))
        
        If ($Option -eq "r") { 
            Write-Host "Reverting to initial search folder scope for Compliance search $searchname"
            Set-ComplianceSearch $searchname -ContentMatchQuery $OldContentMatchQuery
          Do {
                $search = Get-ComplianceSearch $searchname
                Write-Host " > current Search status is $($search.Status) and search job progress: $($search.JobProgress)%"
                Start-Sleep -Seconds 5
             } While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))
            Write-Host "Done"
                            }
        Else {Write-Host "Compliance search $searchname remains configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
        $ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch
        $latestquery = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
    
         [string]$SectionTitle = "Updated Compliance Search detailed information - keep exclusions"
    
         [string]$Description = "Selected $searchname search query has been updated to keep the exclusions: $latestquery"
    
         [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Yellow" -Description $Description -DataType "String" -EffectiveDataString " "
       
         $null = $TheObjectToConvertToHTML.Add($SectionHTML)
    
    }
    }  
    
    Clear-Host
    #$Workloads = "exo", "SCC"
       
    #Connect-O365PS $Workloads
    
    ### Scenario A.1. - Single mailbox scenario - Delete search items that are accessible to user
    
    # Select the search from existing searches
    
    $allSearches = Get-ComplianceSearch
    Write-Host -ForegroundColor Yellow "Please select the single mailbox search for which you wish to delete the found items:"
     
    [string]$SelectedSearch = ($allSearches | Select-Object name | Out-GridView -OutputMode single -Title "Select one search").Name
    
    $ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch
    
    $initialitems = $ComplianceSearch.Items
    
    $searchname = $ComplianceSearch.Name
    
    $query = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
    
    $contentsize = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].contentsize
    
    Write-Host "Found " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems " -NoNewline; Write-Host -ForegroundColor white "items for compliance Search " -NoNewline; Write-Host -ForegroundColor yellow "$searchname"
    
    # Check if search 'locations' contains only one '@' symbol, else prompt 'Not single mailbox search, please select a search for a single mailbox' and loop at beginning.
        If ($compliancesearch.ExchangeLocation.count -ne 1) {
            write-host ""
            Write-Log
            Read-Key
            Start-O365TroubleshootersMenu
        }
    
    # Confirmation input
    
    Write-Host -ForegroundColor Magenta "Are you sure you want to delete $initialitems items found by the selected $searchname Compliance Search?"
    
    [string]$Option = read-host "Type 'yes' to confirm"
    $option = $option.ToLower()
    If ($Option -ne "yes") { exit }
    
        
        $searchdetails = New-Object -TypeName [PSCustomObject]@{
            Name = $searchname,
            items = $initialitems,
            searchquery = $query,
            size = $contentsize
        }
        
        $TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"
    
        $ts= get-date -Format yyyyMMdd_HHmmss
    
        $ExportPath = "$global:WSPath\ComplianceSearch_bulkdelete_$ts"
        
        mkdir $ExportPath -Force |out-null
    
         $searchdetails | export-csv -Path $ExportPath\Selected_search_details.csv -NoTypeInformation 
        
         [string]$SectionTitle = "Compliance Search detailed information"
    
         [string]$Description = "Selection from Get-ComplianceSearch for the selected search"
    
         [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $searchdetails -tabletype List 
         # [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString " "
         # [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $searchdetails -tabletype Table 
         $null = $TheObjectToConvertToHTML.Add($SectionHTML)
    
        # $items = (Get-ComplianceSearch $searchname).Items
        # Write-Progress -Activity "Purging $initialitems items" -Status "$items items left" -PercentComplete ($Iteration / $Iterations * 100)
        
        
    
    # Identifying 'Recoverable Items', 'Purges' and 'Versions' folders - this part is taken from article:  
    
    Write-Host "Identifying 'Recoverable Items', 'Purges' and 'Versions' folders"
    
    [string]$mbx = $compliancesearch.exchangelocation # to add check to confirm there is a single mailbox in locations
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
        $folderIdBytes | Select-Object -skip 23 -First 24 | % { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
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
    
    Set-ComplianceSearch $searchname -ContentMatchQuery $NewContentMatchQuery
      Do {
            $search = Get-ComplianceSearch $searchname
            Write-Host " > current Search status is $($search.Status) and search job progress: $($search.JobProgress)%"
            Start-Sleep -Seconds 5
         } While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))
        
    $Iterations = [math]::Ceiling($initialitems / 10)
    $iteration = 0
    
    # Starting bulk deletion
    
    write-host "Deleting " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems" -NoNewline; Write-Host -ForegroundColor white " items in " -NoNewline; Write-Host -ForegroundColor yellow "$iterations" -NoNewline; Write-Host -ForegroundColor white " batches of 10 items each, due to Compliance Search Action limit"
    
    DO {
        $iteration++
    
        #Write-Host "Refreshing the Compliance Search"
    
        #Get-ComplianceSearch $searchname | Start-ComplianceSearch
        $PurgeAction = Get-ComplianceSearchAction | Where-Object { ($_.Name -match "$searchname") -and ($_.Name -match "_Purge") }
        If ($PurgeAction) {
            foreach ($p in $PurgeAction) { Remove-ComplianceSearchAction ($p).name -Confirm:$False | Out-Null }
        }
                            
        Start-Sleep -Seconds 5
        #$items = (Get-ComplianceSearch $searchname).Items
        #Write-Progress -Activity "Purging $initialitems items" -Status "$items items left" -PercentComplete ($Iteration / $Iterations * 100)
    
        Write-Host -ForegroundColor Yellow "Batch no. [$Iteration / $Iterations]"
        [boolean]$repeat = $True
        $i = 1
        while ($repeat) {
            
            try {
                New-ComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete -Confirm:$false | Out-Null
                $repeat=$false
            }
            catch {
                Start-Sleep -Seconds 5
                Write-Log -function "Start-complianceSearchBulkDelete" -step  "Create new Compliance Search Action" -Description "$($PSItem.Exception.Message)"
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
        
        Do {
            $PurgeAction = Get-ComplianceSearchAction -Identity "$searchname`_Purge"
            Write-Host " > current batch Purge Action status: $($PurgeAction.Status)"
            Start-Sleep -Seconds 5
    
        } While ($PurgeAction.Status -ne 'Completed')
    
    
    } While ($iteration -ne $Iterations)
    
      Write-Host -ForegroundColor Yellow "Finished deleting $initialitems items found by Compliance Search $searchname!"
    
      Remove-ComplianceSearchAction ($PurgeAction).name -Confirm:$False | Out-Null
    
    &$cleanup
    
    <#
    Write-Host "Currently the search is configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
    Write-Host -ForegroundColor Magenta "How do you prefer to keep the search?"
    Write-Host "Type 'K' to KEEP the exclusion. If you re-run the search, it will NOT find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    Write-Host
    Write-Host "Type 'R' to Remove the exclusion. If you re-run the search, it will find the initial items, but they are available under 'Purges' mailbox folder, governed by the mailbox retention settings and holds applied."
    
    Do {
    [string]$option = read-host "Please type one of the above mentioned options"
    $option = $option.ToLower()
    } Until (($option -eq "k") -or ($option -eq "r"))
        
        If ($Option -eq "r") { 
            Write-Host "Reverting to initial search folder scope for Compliance search $searchname"
            Set-ComplianceSearch $searchname -ContentMatchQuery $OldContentMatchQuery
          Do {
                $search = Get-ComplianceSearch $searchname
                Write-Host " > current Search status is $($search.Status) and search job progress: $($search.JobProgress)%"
                Start-Sleep -Seconds 5
             } While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))
            Write-Host "Done"
                            }
        Else {Write-Host "Compliance search $searchname remains configured to exclude the 'Recoverable Items', 'Purges' and 'Versions' folders from its scope."
        $ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch
        $latestquery = ($ComplianceSearch.searchstatistics | ConvertFrom-Json).exchangebinding.queries[0].query
    
         [string]$SectionTitle = "Updated Compliance Search detailed information - keep exclusions"
    
         [string]$Description = "Selected $searchname search query has been updated to keep the exclusions: $latestquery"
    
         [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Yellow" -Description $Description -DataType "String" -EffectiveDataString " "
       
         $null = $TheObjectToConvertToHTML.Add($SectionHTML)
    
    }
    #>  
    #Build HTML report out of the previous HTML sections
    
    [string]$FilePath = $ExportPath + "\Selected_search_report.html"
    
    Export-ReportToHTML -FilePath $FilePath -PageTitle "Compliance Search Bulk Delete Action execution report" -ReportTitle "Compliance Search Bulk Delete Action execution report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
    
    #Ask end-user for opening the HTMl report
    
    $OpenHTMLfile = Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
    
    if ($OpenHTMLfile.ToLower() -like "*y*") {
    
        Write-Host "Opening report...." -ForegroundColor Cyan
    
        Start-Process $FilePath
    
    }
    
    #endregion ResultReport
    
    # Print location where the data was exported
    
    Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
    
    Read-Key
    
    Start-O365TroubleshootersMenu
    
    