<#
        .SYNOPSIS
        Decode Microsoft Defender for Office 365 Safe Links to show original URL 

        .DESCRIPTION
        Provide Microsoft Defender for Office 365 Safe Links and export in a HTML format the original URL
        Can be executed on multiple encoded URL and in the end all decoded URLs can be seen the the HTML output

        .EXAMPLE
        Provide the re-written URL:
        https://nam06.safelinks.protection.outlook.com/?url=http://www.contoso.com/&data=04|01|user1@contoso.com|83ffsdfa384443fadq342743b|72f988fasdfa4d011db47|1|0|6376688415|Unknown|TWFpbGZMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwfadsfaCI6Mn0=|1000&sdata=qOwctqh5fadfaai/tglS4avTxToy67X4M8fadsfasaA=&reserved=0
        
        .LINK
        Online documentation: https://answers.microsoft.com/

    #>
    Clear-Host


### Scenario A.1. - Single mailbox scenario - Items are accessible to user (items are not under 'Recoverable Items' folder)

# select the search from existing

$allSearches = Get-ComplianceSearch 
[string]$SelectedSearch = ($allSearches | select name |Out-GridView -PassThru -Title "Select one search").Name

$ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch

$ComplianceSearch

$items = $ComplianceSearch.Items

$searchname = $ComplianceSearch.Name

Write-Host "Found $items items for compliance Search $searchname"

# Identifying 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders 

Write-Host "Identifying 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders"

[string]$mbx = $compliancesearch.exchangelocation
$folderQueries = @()
   $folderStatistics = Get-MailboxFolderStatistics $mbx | where-object {($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Deletions") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions")}
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
   $DeletionsFolder = $folderQueries.folderquery[1]
   $PurgesFolder = $folderQueries.folderquery[2]
   $VersionsFolder = $folderQueries.folderquery[3]

# Adjusting the search scope to exclude 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders 

Write-Host "Adjusting the search scope to exclude 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders"

[string]$OldContentMatchQuery = $ComplianceSearch.ContentMatchQuery
[string]$NewContentMatchQuery = [string]$OldContentMatchQuery + "(NOT (($RecoverableItemsFolder) OR ($DeletionsFolder) OR ($PurgesFolder) OR ($VersionsFolder)))"

Set-ComplianceSearch $searchname -ContentMatchQuery $NewContentMatchQuery

 $Iterations = [math]::Ceiling($items / 10)

foreach ($Iteration in @(1..$Iterations)) {

Write-Host "Refreshing the Compliance Search"

Get-ComplianceSearch $searchname | Start-ComplianceSearch

Do {

    Start-Sleep -Seconds 2
    $search = Get-ComplianceSearch $searchname
    Write-Host "Current Search status: $($search.Status) and search job progress: $($search.JobProgress)"

} While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))

Write-Progress -Activity "Purging items" -Status "$items items left" -PercentComplete ($Iteration / $Iterations * 100)

Write-Host "iteration no. [$Iteration / $Iterations]"

New-ComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

Do {

    Start-Sleep -Seconds 2

    $PurgeAction = Get-ComplianceSearchAction -Identity "$searchname`_Purge"

    Write-Host " > current iteration's Purge Action status: $($PurgeAction.Status)"

} While ($PurgeAction.Status -ne 'Completed')
Remove-ComplianceSearchAction ($PurgeAction).name -Confirm:$False | Out-Null
}



### Scenario A.2. - Single mailbox scenario - Purge items from 'Recoverable Items' folder and subfolders


# select the search from existing

$allSearches = Get-ComplianceSearch 
[string]$SelectedSearch = ($allSearches | select name |Out-GridView -PassThru -Title "Select one search").Name

$ComplianceSearch = Get-ComplianceSearch -Identity $SelectedSearch

$ComplianceSearch

$items = $ComplianceSearch.Items

$searchname = $ComplianceSearch.Name

Write-Host "Found $items items for compliance Search $searchname"

# Identifying 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders 

Write-Host "Identifying 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders"

[string]$mbx = $compliancesearch.exchangelocation
$folderQueries = @()
   $folderStatistics = Get-MailboxFolderStatistics $mbx | where-object {($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Deletions") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions")}
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
   $DeletionsFolder = $folderQueries.folderquery[1]
   $PurgesFolder = $folderQueries.folderquery[2]
   $VersionsFolder = $folderQueries.folderquery[3]

# Adjusting the search scope to exclude 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders 

Write-Host "Adjusting the search scope to target only the 'Recoverable Items', 'Deletions', 'Purges' and 'Versions' folders"

[string]$OldContentMatchQuery = $ComplianceSearch.ContentMatchQuery
[string]$NewContentMatchQuery = [string]$OldContentMatchQuery + "(($RecoverableItemsFolder) OR ($DeletionsFolder) OR ($PurgesFolder) OR ($VersionsFolder))"

Set-ComplianceSearch $searchname -ContentMatchQuery $NewContentMatchQuery

 $Iterations = [math]::Ceiling($items / 10)

foreach ($Iteration in @(1..$Iterations)) {

Write-Host "Refreshing the Compliance Search"

Get-ComplianceSearch $searchname | Start-ComplianceSearch

Do {

    Start-Sleep -Seconds 2
    $search = Get-ComplianceSearch $searchname
    Write-Host "Current Search status: $($search.Status) and search job progress: $($search.JobProgress)"

} While (($search.Status -ne 'Completed') -and ($search.JobProgress -ne '100'))

Write-Progress -Activity "Purging items" -Status "$items items left" -PercentComplete ($Iteration / $Iterations * 100)

Write-Host "iteration no. [$Iteration / $Iterations]"

New-ComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

Do {

    Start-Sleep -Seconds 2

    $PurgeAction = Get-ComplianceSearchAction -Identity "$searchname`_Purge"

    Write-Host " > current iteration's Purge Action status: $($PurgeAction.Status)"

} While ($PurgeAction.Status -ne 'Completed')
Remove-ComplianceSearchAction ($PurgeAction).name -Confirm:$False | Out-Null
}
