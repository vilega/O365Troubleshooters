<# 1st requirement install the module O365 TS
Import-Module C:\Users\alexaca\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>

<#

        .SYNOPSIS

        Get a report of mailbox folder permissions for one or more mailboxes



        .DESCRIPTION

        Get a report of mailbox folder permissions for one or more mailboxes...



        .EXAMPLE

        If we check a mailbox for folder permissions, we can find out if any of the default folders have modified default permissions and someone else has access to the contents of those folders

        

        .LINK

        Online documentation: https://aka.ms/O365Troubleshooters/CheckMailboxFolderPermissions



    #>

function Get-UserMailboxFolderPermissions {

    param(
        [System.Collections.ArrayList]$MBXs,
        [bool]$isDefaultFolder)
    
    
    $rights = New-Object -TypeName "System.Collections.ArrayList"
    $foldersForAllMbx = @()
    
    foreach ($MBX in $MBXs) {
        #Extracting the Name attribute for the mailbox
        $alias = (Get-Mailbox $MBX).Name.ToString()
        #Extracting the Primary SMTP Address for the mailbox
        $SMTP = (Get-Mailbox $MBX).PrimarySMTPAddress.ToString()
        #Getting all mailbox folders list
        if ($isDefaultFolder -eq $true) {
            [System.Collections.ArrayList]$folders = get-mailbox $MBX | Get-MailboxFolderStatistics | Where-Object FolderType -ne "User Created" | select Identity, @{Name = 'Alias'; Expression = { $alias } } , @{Name = 'SMTP'; Expression = { $SMTP } } 
        }
        else {
            [System.Collections.ArrayList]$folders = get-mailbox $MBX | Get-MailboxFolderStatistics | select-object Identity, @{Name = 'Alias'; Expression = { $alias } } , @{Name = 'SMTP'; Expression = { $SMTP } } 
        }
        $foldersForAllMbx += $folders

        #With below 2 command lines I am attempting to get the Top of Information Store folder permission as well in the mailbox.
        $MBRightsRoot = Get-MailboxFolderPermission -Identity "$MBX" -ErrorAction Stop
        $MBRightsRoot = $MBRightsRoot | Select-Object FolderName, User, AccessRights, @{Name = 'SMTP'; Expression = { $SMTP } }
        $null = $rights.Add($MBRightsRoot)

    }
    write-log -Function "Start-CheckMBXFolderPermissions" -Step "Get-UserMailboxFolderPermissions" -Description "Preparing folder name format for input on Get-MailboxFolderPermission cmdlet"

    #Adjusting the folder Identity values obtained by previous command, to comply with the Get-MailboxFolderPermission cmdlet required format. Getting the folder permissions as well.
    foreach ($folder in $foldersForAllMbx) {
        $foldername = $folder.Identity.ToString().Replace([char]63743, "/").Replace($folder.alias, $folder.SMTP + ":")
        try {
            $MBrights = Get-MailboxFolderPermission -Identity "$foldername" -ErrorAction Stop
            [System.Collections.ArrayList]$MBrights = $MBrights | Select-Object FolderName, User, AccessRights, @{Name = 'SMTP'; Expression = { $folder.SMTP } }
                
            $null = $rights.Add($MBrights)
           
        }
        Catch {
            #TODO: Need to implement error handling in a future revision}
        }
        write-log -Function "Start-CheckMBXFolderPermissions" -Step "Get-UserMailboxFolderPermissions" -Description "Getting permissions for the list of mailbox folders"
    }
    return ($rights)
} 
# connect
Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads

# logging
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Start-CheckMBXFolderPermissions" -Step $CurrentProperty -Description $CurrentDescription 

$ts = get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\MailboxDiagnosticLogs_$ts"
mkdir $ExportPath -Force | out-null


$allMBX = Get-ExoMailbox -Filter "RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'SharedMailbox'" | select DisplayName, PrimarySmtpAddress, UserPrincipalName
Write-Host "Warning: Please keep in mind that the more mailboxes are selected, this will affect the performance of the script" -ForegroundColor Yellow
$choice = Read-Host "Please select the mailboxes that need to be checked (press Enter to display the list of mailboxes)"
#$allMBXInitialCount = $allMBX.Count#
[Array]$allMBX = ($allMBX | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName | Out-GridView -PassThru -Title "Select one or more..").PrimarySmtpAddress
 
$allMBXSelectedCount = $allMBX.Count
    
If ($allMBXSelectedCount -eq 0) {
    write-log -Function "Start-CheckMBXFolderPermissions" -Step "SelectedMailboxes" -Description "Fail"  
    Write-Host "You have made no selection, we will return to the main menu!"
    Read-Key
    Start-O365TroubleshootersMenu
}
write-log -Function "Start-CheckMBXFolderPermissions" -Step "SelectedMailboxes" -Description "Success"

Write-Host "Warning: Depending on the number of mailboxes selected, running the script to check all folders, might give a timeout" -ForegroundColor Yellow
#$choice = Read-Host "Do you want to check all folders or only default ones? Input '1' for 'All folders' or '2' for 'Default folders'"
$choice = Get-Choice -Options 'All Folders', 'Default Folders'   
if ($choice -eq "d") {
    $isDefaultFolder = $true
}
elseif ($choice -eq "a") {
    $isDefaultFolder = $false
}
write-log -Function "Start-CheckMBXFolderPermissions" -Step "Chose only default folders" -Description $isDefaultFolder   
write-log -Function "Start-CheckMBXFolderPermissions" -Step "Get-UserMailboxFolderPermissions" -Description "Calling function"

$rights = Get-UserMailboxFolderPermissions -MBXs $allMBX -isDefaultFolder $isDefaultFolder

write-log -Function "Start-CheckMBXFolderPermissions" -Step "Get-UserMailboxFolderPermissions" -Description "Returning from function"

$ExportRights = $rights | ForEach-Object { $_ }
write-log -Function "Start-CheckMBXFolderPermissions" -Step "Export CSV with mailbox folder permissions for selected users" -Description "Success"
$ExportRights | Export-Csv $ExportPath\Mailbox_Folder_Permissions_$ts.csv -NoTypeInformation


#Create the collection of sections of HTML

$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"

[string]$SectionTitle = "Information"
[string]$Description = "We will highlight in red all the users which have folders where the Default user permission is different than None. This can impact security as this type permission allows all users in the organization to access that user's folder. Please review!"
[PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString ""
$null = $TheObjectToConvertToHTML.Add($SectionHTML)



foreach ($mailbox in $allMBX) {
    [string]$SectionTitle = "Information for the following mailbox: $($mailbox)"
   
    $ExportRightsCurrentMbx = $ExportRights | Where-Object SMTP -eq  $mailbox | Select-Object * -ExcludeProperty SMTP

    $defaultuser = $true

    foreach ($currentrights in $ExportRightsCurrentMbx) {

        if (($currentrights.User.DisplayName -eq "Default") -and !($currentrights.AccessRights -eq "None") -and !($currentrights.AccessRights -eq "AvailabilityOnly")) {
            $defaultuser = $false

        }


    }

    if ($isDefaultFolder -eq $true) {
        [string]$Description = "We take a look only at the mailbox's default folder permissions"
    }

    elseif ($isDefaultFolder -eq $false) {
        [string]$Description = "We take a look at all mailbox's folder permissions"   
    }

    if ($defaultuser -eq $True) {
        [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -TableType "Table" -EffectiveDataArrayList $ExportRightsCurrentMbx
    }
    else {
        [PSCustomObject]$SectionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -TableType "Table" -EffectiveDataArrayList $ExportRightsCurrentMbx
    }


    $null = $TheObjectToConvertToHTML.Add($SectionHTML)

}



#Build HTML report out of the previous HTML sections

[string]$FilePath = $ExportPath + "\MailboxFolderPermissions.html"

Export-ReportToHTML -FilePath $FilePath -PageTitle "Check Mailbox Folder Permissions" -ReportTitle "Check Mailbox Folder Permissions" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
write-log -Function "Start-CheckMBXFolderPermissions" -Step "Export HTML report" -Description "Success"

#Ask end-user for opening the HTMl report

$OpenHTMLfile = Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"

if ($OpenHTMLfile.ToLower() -like "*y*") {

    Write-Host "Opening report...." -ForegroundColor Cyan

    Start-Process $FilePath

}

#endregion ResultReport

   

# Print location where the data was exported

Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
#>

Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu