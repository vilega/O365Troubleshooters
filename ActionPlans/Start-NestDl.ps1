Function Show-AzureADGroupMembersIncludedNested {
    param([string]$PrimarySmtpAddress)
    [System.Collections.ArrayList]$AllMembers =@()
    $members = Get-AzureADGroup -SearchString $PrimarySmtpAddress.Split("@")[0]| foreach {Get-AzureADGroupMember -ObjectId $_.objectid} | Select-Object MailNickName, Mail, ObjectId, ObjectType
    foreach ($member in $members) {
        $entry = New-Object -TypeName psobject
        $entry | Add-Member -MemberType NoteProperty -Name ParentGroup -Value $PrimarySmtpAddress
        $entry | Add-Member -MemberType NoteProperty -Name Name -Value $member.MailNickName
        $entry | Add-Member -MemberType NoteProperty -Name Mail -Value $member.Mail
        $entry | Add-Member -MemberType NoteProperty -Name ObjectId -Value $member.ObjectId
        $entry | Add-Member -MemberType NoteProperty -Name ObjectType -Value $member.ObjectType
        $AllMembers.add($entry) |Out-Null
        if ($member.ObjectType -like "Group") {
            Show-AzureADGroupMembersIncludedNested -PrimarySmtpAddress $member.Mail
        }
    }
    return $AllMembers
    $AllMembers|Export-Csv -Append -NoTypeInformation AllMembers.csv
}
Function Show-DLMembersIncludedNested {
    param([string]$PrimarySmtpAddress)
    [System.Collections.ArrayList]$AllMembers =@()
    #validate Dl RecipientTypeDetails
    $Dl=Get-Recipient $PrimarySmtpAddress | Select-Object Name, PrimarySmtpAddress,RecipientType ,RecipientTypeDetails
    if($Dl.RecipientTypeDetails -like "GroupMailbox")
    {
    $members = Get-UnifiedGroupLinks -LinkType member -Identity $PrimarySmtpAddress |Select-Object Name,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,ExternalDirectoryObjectId
    }
    elseif($Dl.RecipientTypeDetails -like "*Group" -or $Dl.RecipientTypeDetails -like "RoomList")
    {
    $members = Get-DistributionGroupMember -Identity $PrimarySmtpAddress | Select-Object Name,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,ExternalDirectoryObjectId
    }
    foreach ($member in $members) {
        $entry = New-Object -TypeName psobject
        $entry | Add-Member -MemberType NoteProperty -Name ParentDLSMTP -Value $PrimarySmtpAddress
        $entry | Add-Member -MemberType NoteProperty -Name Name -Value $member.Name
        $entry | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $member.PrimarySmtpAddress
        $entry | Add-Member -MemberType NoteProperty -Name ExternalDirectoryObjectId -Value $member.ExternalDirectoryObjectId
        $entry | Add-Member -MemberType NoteProperty -Name RecipientTypeDetails -Value $member.RecipientTypeDetails
        $AllMembers.add($entry) |Out-Null
        if ($member.RecipientType -like "*group") {
            Show-DLMembersIncludedNested -PrimarySmtpAddress $member.PrimarySmtpAddress
        }
    }
    return $AllMembers
}