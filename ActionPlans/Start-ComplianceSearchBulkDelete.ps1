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
    Clear-Host


### Scenario A.1. - Single mailbox scenario - Items are accessible to user (items are not under 'Recoverable Items' folder)

    # Select the search from existing searches

$allSearches = Get-ComplianceSearch
Write-Host -ForegroundColor Yellow "Please select the search for which you wish to delete the found items:"
 
[string]$SelectedSearch = ($allSearches | select name |Out-GridView -OutputMode single -Title "Select one search").Name

$ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch

$initialitems = $ComplianceSearch.Items

$searchname = $ComplianceSearch.Name

Write-Host "Found " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems " -NoNewline; Write-Host -ForegroundColor white "items for compliance Search " -NoNewline; Write-Host -ForegroundColor yellow "$searchname"


    # Confirmation input - to be replaced with 'Choice' function

    Write-Host -ForegroundColor Magenta "Are you sure you want to delete $initialitems items found by the selected $searchname Compliance Search?"
    [string]$Option = read-host "Type 'yes' to confirm"

    If ($Option -ne "yes") {exit}

    # Identifying 'Recoverable Items', 'Purges' and 'Versions' folders 

Write-Host "Identifying 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders"

[string]$mbx = $compliancesearch.exchangelocation # to add check to confirm there is a single mailbox in locations
$folderQueries = @()
   $folderStatistics = Get-MailboxFolderStatistics $mbx | where-object {($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions")}
   foreach ($folderStatistic in $folderStatistics)
   {
       $folderId = $folderStatistic.FolderId;
       $folderPath = $folderStatistic.FolderPath;
       $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
       $nibbler= $encoding.GetBytes("0123456789ABCDEF");
       $folderIdBytes = [Convert]::FromBase64String($folderId);
       $indexIdBytes = New-Object byte[] 48;
       $indexIdIdx=0;
       $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
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
$Iterations = [math]::Ceiling($initialitems / 10)
$iteration = 0

    # Starting bulk deletion

write-host "Deleting " -NoNewline; Write-Host -ForegroundColor Yellow "$initialitems" -NoNewline; Write-Host -ForegroundColor white " items in " -NoNewline; Write-Host -ForegroundColor yellow "$iterations" -NoNewline; Write-Host -ForegroundColor white " batches of 10 items each, due to Compliance Search Action limit"

DO {

   $iteration++

   Write-Host "Refreshing the Compliance Search"

        Get-ComplianceSearch $searchname | Start-ComplianceSearch
        $PurgeAction = Get-ComplianceSearchAction | Where-Object {($_.Name -match "$searchname") -and ($_.Name -match "_Purge")}
        If ($PurgeAction) {
            foreach ($p in $PurgeAction) {Remove-ComplianceSearchAction ($p).name -Confirm:$False | Out-Null}
            }
                        
            Do {
                $search = Get-ComplianceSearch $searchname
                Write-Host " > current Search status is $($search.Status) and search job progress: $($search.JobProgress)%"
                Start-Sleep -Seconds 5
            } While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))

        Start-Sleep -Seconds 10
        $items = (Get-ComplianceSearch $searchname).Items
        Write-Progress -Activity "Purging $initialitems items" -Status "$items items left" -PercentComplete ($Iteration / $Iterations * 100)

        Write-Host -ForegroundColor Yellow "Batch no. [$Iteration / $Iterations]"
        New-ComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

            Do {
                $PurgeAction = Get-ComplianceSearchAction -Identity "$searchname`_Purge"
                Write-Host " > current batch Purge Action status: $($PurgeAction.Status)"
                Start-Sleep -Seconds 5

            } While ($PurgeAction.Status -ne 'Completed')

        Remove-ComplianceSearchAction ($PurgeAction).name -Confirm:$False | Out-Null

} While ($iteration -ne $Iterations)

Write-Host -ForegroundColor Yellow "Finished deleting $initialitems items found by Compliance Search $searchname!"

