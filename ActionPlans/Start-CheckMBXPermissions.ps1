
<# 
Import-Module C:\Users\alexaca\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
Start-O365TroubleshootersMenu
#>


function Get-AllNoneUserDefaultMailboxFolderPermissions {

    #Extracting the Name attribute for the mailbox
    
    $alias=(Get-Mailbox $MBX).Name.ToString()
    
    #Extracting the Primary SMTP Address for the mailbox
    
    $SMTP =(Get-Mailbox $MBX).PrimarySMTPAddress.ToString()
    
    #Getting the default mailbox folders list
    
    $folders=get-mailbox $MBX | Get-MailboxFolderStatistics | ? {($_.FolderType -eq "Inbox") -or ($_.FolderType -eq "Archive") -or ($_.FolderType -eq "Drafts") -or ($_.FolderType -eq "Outbox") -or ($_.FolderType -eq "Calendar") -or ($_.FolderType -eq "SentItems") -or ($_.FolderType -eq "Contacts") -or ($_.FolderType -eq "Tasks") -or ($_.FolderType -eq "Notes") -or ($_.FolderType -eq "DeletedItems") -or ($_.FolderType -eq "JunkEmail")} | select Identity
    
    $rights = @()
    
    #Adjusting the folder Identity values obtained by previous command, to comply with the Get-MailboxFolderPermission cmdlet required format. Getting the folder permissions as well.
    
    foreach ($folder in $folders) {
    
    $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($alias,$SMTP + ":")
    
    try
    
    {
    
    $MBrights = Get-MailboxFolderPermission -Identity "$foldername" -ErrorAction Stop | ? {($_.User -eq "Default") -and ($_.AccessRights -ne "None")}
    
    $MBrights =$MBrights | Select FolderName,User, AccessRights, @{Name = 'SMTP'; Expression = {$SMTP}}
    
    #With below 2 command lines I am attempting to get the Top of Information Store folder permission as well in the mailbox.
    
    $MBRightsRoot = Get-MailboxFolderPermission -Identity "$MBX" -ErrorAction Stop | ? {($_.User -eq "Default") -and ($_.AccessRights -ne "None")}
    
    $MBRightsRoot = $MBRightsRoot | Select FolderName,User, AccessRights, @{Name = 'SMTP'; Expression = {$SMTP}}
    
    $rights += $MBrights
    }
    
    Catch {}
    
    }
    
    return ($rights + $MBRightsRoot)
    
    }



function Get-AllDefaultUserMailboxFolderPermissions {

    #Extracting the Name attribute for the mailbox
    
    $alias=(Get-Mailbox $MBX).Name.ToString()
    
    #Extracting the Primary SMTP Address for the mailbox
    
    $SMTP =(Get-Mailbox $MBX).PrimarySMTPAddress.ToString()
    
    #Getting the default mailbox folders list
    
    $folders=get-mailbox $MBX | Get-MailboxFolderStatistics | ? {($_.FolderType -eq "Inbox") -or ($_.FolderType -eq "Archive") -or ($_.FolderType -eq "Drafts") -or ($_.FolderType -eq "Outbox") -or ($_.FolderType -eq "Calendar") -or ($_.FolderType -eq "SentItems") -or ($_.FolderType -eq "Contacts") -or ($_.FolderType -eq "Tasks") -or ($_.FolderType -eq "Notes") -or ($_.FolderType -eq "DeletedItems") -or ($_.FolderType -eq "JunkEmail")} | select Identity
    
    $rights = @()
    
    #Adjusting the folder Identity values obtained by previous command, to comply with the Get-MailboxFolderPermission cmdlet required format. Getting the folder permissions as well.
    
    foreach ($folder in $folders) {
    
    $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($alias,$SMTP + ":")
    
    try
    
    {
    
    $MBrights = Get-MailboxFolderPermission -Identity "$foldername" -ErrorAction Stop
    
    $MBrights =$MBrights | Select FolderName,User, AccessRights, @{Name = 'SMTP'; Expression = {$SMTP}}
    
    #With below 2 command lines I am attempting to get the Top of Information Store folder permission as well in the mailbox.
    
    $MBRightsRoot = Get-MailboxFolderPermission -Identity "$MBX" -ErrorAction Stop
    
    $MBRightsRoot = $MBRightsRoot | Select FolderName,User, AccessRights, @{Name = 'SMTP'; Expression = {$SMTP}}
    
    $rights += $MBrights
    }
    
    Catch {}
    
    }
    
    return ($rights + $MBRightsRoot)
    
    }


# connect
Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads

# logging
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 

$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\MailboxDiagnosticLogs_$ts"
mkdir $ExportPath -Force |out-null


#Gathering a list of the User Principal Name attribute for the specified recipient type

Write-Host "======================" -ForegroundColor Gray

Write-Host "SharedMailbox" -ForegroundColor Green

Write-Host "UserMailbox" -ForegroundColor Yellow

Write-Host "RoomMailbox" -ForegroundColor Cyan

Write-Host "EquipmentMailbox" -ForegroundColor Magenta

Write-Host "======================" -ForegroundColor Gray

$Type=Read-Host -Prompt "Please type the RecipientTypeDetails for which you wish to get the information, available values are shown above"

$UPN = Get-Mailbox -RecipientTypeDetails $Type -ResultSize Unlimited | select UserPrincipalName

$UPN=($UPN).UserPrincipalName

foreach ($MBX in $UPN) {


#Export a CSV file for each mailbox, that contains the default folders, their associated permissions and the Primary SMTP address of the mailbox in question.


Get-AllDefaultUserMailboxFolderPermissions | Export-Csv -Path "$(($MBX.ToString()))_MailboxFolderPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture


}


foreach ($MBX in $UPN) {


    #Export a CSV file for each mailbox, that contains the default folders where the Default user has AccessRights set to other than None, their associated permissions and the Primary SMTP address of the mailbox in question.
    
    
    Get-AllNoneDefaultUserMailboxFolderPermissions | Export-Csv -Path "$(($MBX.ToString()))_NoneDefaultMailboxFolderPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
    
 
    
    }


Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu