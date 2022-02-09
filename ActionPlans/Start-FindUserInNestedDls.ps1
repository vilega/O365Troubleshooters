<# 1st requirement install the module O365 TS
Import-Module C:\Users\haembab\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
#>
Function Show-DlAADEXOSyncDiscrepancies{
    param([System.Collections.ArrayList]$AADmbrresults,[System.Collections.ArrayList]$EXOmbrresults)
    if ($AADmbrresults.count -ne $EXOmbrresults.count)
    {
        $AADmbrresults=$AADmbrresults |Group-Object parentgroup
        $EXOmbrresults=$EXOmbrresults |Group-Object ParentDLSMTP
        [System.Collections.ArrayList]$NotfoundinEXO =@()
        foreach($AADmbrresult in $AADmbrresults)
        {
            foreach($EXOmbrresult in $EXOmbrresults)
            {
                #ParentDLmatch
                if($AADmbrresult.Name -eq $EXOmbrresult.Name)
                {
                #check count
                    if($AADmbrresult.Count -ne $EXOmbrresult.Count)
                    {
                    #check discrepancy
                        foreach($AADmbr in $AADmbrresult.Group)
                        {
                            $counter=1
                            foreach($EXOmbr in $($EXOmbrresult.Group))
                            {
                                if($AADmbr.ObjectId -eq $EXOmbr.ExternalDirectoryObjectId)
                                {
                                #match found reseetting counter
                                $counter=0
                                }
                                elseif($AADmbr.ObjectId -ne $EXOmbr.ExternalDirectoryObjectId -and $counter -ge $($EXOmbrresult.Group).count)
                                {
                                    $NotfoundinEXO=$NotfoundinEXO+$AADmbr
                                }
                            $counter++
                            }                     
                        }
                    }
                }
            }    
        }
    }
    return $NotfoundinEXO
}
Clear-Host
#TODO check with victor for AzureADPreview import/install issue
$Workloads = "exo,AzureAdPreview"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
Clear-Host
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\DlAADEXOSyncDiscrepancies_$ts"
mkdir $ExportPath -Force |Out-Null
$PrimarySmtpAddress=Get-ValidEmailAddress("Email address of the Distribution Group ")
[System.Collections.ArrayList]$AADmbrresults =@()
[System.Collections.ArrayList]$EXOmbrresults =@()
#. $script:modulePath\ActionPlans\Start-NestDl.ps1
$AADmbrresults=Show-AzureADGroupMembersIncludedNested -PrimarySmtpAddress $PrimarySmtpAddress
$EXOmbrresults=Show-DLMembersIncludedNested -PrimarySmtpAddress $PrimarySmtpAddress
$DlAADEXOSyncDiscrepancies=Show-DlAADEXOSyncDiscrepancies($AADmbrresults,$EXOmbrresults)
#HTML Report
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()
[string]$SectionTitle = "Introduction"
[string]$Description = "This report expands NESTED distribution group:$PrimarySmtpAddress members in Azure Active Directory & Exchange online and output members out-of-sync in case found!"
[PSCustomObject]$StartHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please check the next section for more information!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
[string]$SectionTitle = "Distribution group members in AAD"
[string]$Description = "This section lists distribution group members in Azure Active Directory."
[PSCustomObject]$AADmbrresultsHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $AADmbrresults -TableType Table
$null = $TheObjectToConvertToHTML.Add($AADmbrresultsHTML)
[string]$SectionTitle = "Distribution group members in EXO"
[string]$Description = "This section lists distribution group members in Exchange online."
[PSCustomObject]$EXOmbrresultsHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $EXOmbrresults -TableType Table
$null = $TheObjectToConvertToHTML.Add($EXOmbrresultsHTML)
[string]$SectionTitle = "Distribution group members out-of-sync"
[string]$Description = "This section lists distribution group members out-of-sync!"
[PSCustomObject]$DlAADEXOSyncDiscrepanciesHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $DlAADEXOSyncDiscrepancies -TableType Table
$null = $TheObjectToConvertToHTML.Add($DlAADEXOSyncDiscrepanciesHTML)
#region ResultReport
[string]$FilePath = $ExportPath + "\DlAADEXOSyncDiscrepancies.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Dl AAD EXO Sync Discrepancies" -ReportTitle "Dl AAD EXO Sync Discrepancies" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
#Question to ask enduser for opening the HTMl report
$OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile -like "*y*")
{
Write-Host "Opening report...." -ForegroundColor Cyan
Start-Process $FilePath
}
#endregion ResultReport
$AllUsersWithAllRoles | export-csv "$ExportPath\DlAADEXOSyncDiscrepancies_$ts.csv" -NoTypeInformation
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow
Read-Key
Start-O365TroubleshootersMenu