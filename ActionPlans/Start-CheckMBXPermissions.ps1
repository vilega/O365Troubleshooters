
<# 
Import-Module C:\Users\alexaca\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
Start-O365TroubleshootersMenu
#>


function Get-AllDefaultUserMailboxFolderPermissions {

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
            [System.Collections.ArrayList]$folders = get-mailbox $MBX | Get-MailboxFolderStatistics | select Identity, @{Name = 'Alias'; Expression = { $alias } } , @{Name = 'SMTP'; Expression = { $SMTP } } 
        }
        $foldersForAllMbx += $folders
    }
    
    
        
        
    #Adjusting the folder Identity values obtained by previous command, to comply with the Get-MailboxFolderPermission cmdlet required format. Getting the folder permissions as well.
    foreach ($folder in $foldersForAllMbx) {
        $foldername = $folder.Identity.ToString().Replace([char]63743, "/").Replace($folder.alias, $folder.SMTP + ":")
        try {
            $MBrights = Get-MailboxFolderPermission -Identity "$foldername" -ErrorAction Stop
            [System.Collections.ArrayList]$MBrights = $MBrights | Select FolderName, User, AccessRights, @{Name = 'SMTP'; Expression = { $SMTP } }
            #With below 2 command lines I am attempting to get the Top of Information Store folder permission as well in the mailbox.
            $MBRightsRoot = Get-MailboxFolderPermission -Identity "$MBX" -ErrorAction Stop
            [System.Collections.ArrayList]$MBRightsRoot = $MBRightsRoot | Select FolderName, User, AccessRights, @{Name = 'SMTP'; Expression = { $SMTP } }
            foreach ($entry in $MBRightsRoot) {
                $null = $MBrights.Add($entry)
            }
                
            $null = $rights.Add($MBrights)
        }
        Catch {}
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
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

$ts = get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\MailboxDiagnosticLogs_$ts"
mkdir $ExportPath -Force | out-null


$allMBX = Get-ExoMailbox -Filter "RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'SharedMailbox'" | select DisplayName, PrimarySmtpAddress, UserPrincipalName
Write-Host "Warning: Please keep in mind that the more mailboxes are selected, this will affect the performance of the script" -ForegroundColor Yellow
$choice = Read-Host "Please select the mailboxes that need to be checked (press Enter to display the list of mailboxes)"
$allMBXInitialCount = $allMBX.Count
[Array]$allMBX = ($allMBX | select DisplayName, PrimarySmtpAddress, UserPrincipalName | Out-GridView -PassThru -Title "Select one or more..").PrimarySmtpAddress
$allMBXSelectedCount = $allMBX.Count
    
If ($allMBXSelectedCount -eq 0) {
    # go to the menu or get again all mbx
}
    
Write-Host "Warning: Depending on the number of mailboxes selected, running the script to check all folders, might give a timeout" -ForegroundColor Yellow
#$choice = Read-Host "Do you want to check all folders or only default ones? Input '1' for 'All folders' or '2' for 'Default folders'"
$choice = Get-Choice -Options 'All Folders', 'Default Folders'   
if ($choice -eq "d") {
    $isDefaultFolder = $true
}
elseif ($choice -eq "a") {
    $isDefaultFolder = $false
}
    
$rights = Get-AllDefaultUserMailboxFolderPermissions -MBXs $allMBX -isDefaultFolder $isDefaultFolder

$ExportRights = $rights | % { $_ }

$ExportRights | Export-Csv $ExportPath\Mailbox_Folder_Permissions_$ts.csv -NoTypeInformation

#Start-Process     C:\Temp\permissions.csv


Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu