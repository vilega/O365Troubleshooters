<# 1st requirement install the module O365 TS
Import-Module C:\Users\a-haemb\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>
Clear-Host
$DGConditionsmet=0
#region Connecting to EXO & MSOL

$Workloads = "exo","msol"
try {
    Connect-O365PS $Workloads 
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
catch {
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Failure"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
    }
#endregion Connecting to EXO & MSOL

# Create working folder for Groups Diag
# for the HTML file
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\DlToO365GroupUpgradeChecks_$ts"
mkdir $ExportPath -Force |out-null

#region Getting the DG SMTP
$dgsmtp=Get-ValidEmailAddress("Email address of the Distribution Group ")
try {
    $dg=get-DistributionGroup -Identity $dgsmtp -ErrorAction stop
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO Directory"
    $CurrentDescription = "Success"
    write-log -Function "Retrieve Distrubtion Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
catch {
    $CurrentProperty = "Retrieving: $dgsmtp object from EXO"
    $CurrentDescription = "Failure"
    write-log -Function "Retrieve Distrubtion Group Object From EXO Directory" -Step $CurrentProperty -Description $CurrentDescription
}
#endregion Getting the DG SMTP



#Region Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()
$ConditionMemberRestriction=New-Object PSObject
if ($dg.MemberJoinRestriction -eq "Open" -and $dg.MemberDepartRestriction -eq "Open") {
    $ConditionMemberRestriction|Add-Member -NotePropertyName MemberJoinRestriction -NotePropertyValue "Open"
    $ConditionMemberRestriction|Add-Member -NotePropertyName MemberDepartRestriction -NotePropertyValue "Open"
    [string]$SectionTitle = "Validating Distribution Group Member Restriction"
    [string]$Description = "Checking if MemberJoinRestriction & MemberDepartRestriction values are compliant for group to upgrade"
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionMemberRestriction
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)
    #$CurrentProperty = "Checking $dgsmtp if MemberJoinRestriction & MemberDepartRestriction values are compliant for group to upgrade"
    #$CurrentDescription = "Success"
    #write-log -Function "Dl To O365 Groups Checker" -Step $CurrentProperty -Description $CurrentDescription
    # Add the output in html report with green/red accordingly   
    } 
    else {
    $DGConditionsmet++    
    $ConditionMemberRestriction|Add-Member -NotePropertyName MemberJoinRestriction -NotePropertyValue $dg.MemberJoinRestriction
    $ConditionMemberRestriction|Add-Member -NotePropertyName MemberDepartRestriction -NotePropertyValue $dg.MemberDepartRestriction
    [string]$SectionTitle = "Validating Distribution Group Member Restriction"
    [string]$Description = "Checking if MemberJoinRestriction & MemberDepartRestriction values are compliant for group to upgrade"
    [PSCustomObject]$ConditionMemberRestrictionHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $ConditionMemberRestriction
    $null = $TheObjectToConvertToHTML.Add($ConditionMemberRestrictionHTML)
    # Add the output in html report with green/red accordingly       
    }
#endRegion Check if Distribution Group can't be upgraded because Member*Restriction is set to "Closed"

#Region Check if Distribution Group can't be upgraded because it is DirSynced
if ($dg.IsDirSynced -eq $true) {
    $DGConditionsmet++ 
    $CurrentProperty = "Checking $dgsmtp if IsDirSynced value is compliant for group to upgrade"
    $CurrentDescription = "Failure"
    write-log -Function "Dl To O365 Groups Checker" -Step $CurrentProperty -Description $CurrentDescription
    # Add the output in html report with green/red accordingly  
} 
else {
    $CurrentProperty = "Checking $dgsmtp if IsDirSynced value is compliant for group to upgrade"
    $CurrentDescription = "Success"
    write-log -Function "Dl To O365 Groups Checker" -Step $CurrentProperty -Description $CurrentDescription
    # Add the output in html report with green/red accordingly  
}
#endRegion Check if Distribution Group can't be upgraded because it is DirSynced
    
if($DGConditionsmet -gt 0){
    "DG Upgrade Failed"|Out-file -FilePath $ExportPath\result.txt

}
else{
    "DG Upgrade Succeeded"|Out-file -FilePath $ExportPath\result.txt
    
}

#region ResultReport
[string]$FilePath = $ExportPath + "\DistrubtionGroupUpgradeCheck.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Distrubtion Group Upgrade Checker" -ReportTitle "Distrubtion Group Upgrade Checker" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
#endregion ResultReport

    
# End of the Diag
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Start-Sleep -Seconds 3
Read-Key
# Go back to the main menu
Start-O365TroubleshootersMenu
