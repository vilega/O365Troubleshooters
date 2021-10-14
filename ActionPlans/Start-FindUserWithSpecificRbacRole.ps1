<# 1st requirement install the module O365 TS
Import-Module C:\Users\haembab\Documents\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
# 2nd requirement Execute set global variables
Set-GlobalVariables
# 3rd requirement to start the menu
Start-O365TroubleshootersMenu
#>
Clear-Host
$Workloads = "exo"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
Clear-Host
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RbacRoleSpecific_$ts"
mkdir $ExportPath -Force | Out-Null
. $script:modulePath\ActionPlans\Start-RbacTools.ps1
#HTML Report
[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()
$SpecificRoleMembers=Get-SpecificRoleMembers
[string]$SectionTitle = "Introduction"
[String]$article='<a href="https://docs.microsoft.com/en-us/exchange/understanding-role-based-access-control-exchange-2013-help" target="_blank">Understanding role based access control</a>'
[string]$Description = "This report spans all management roles across your enviroment to list users member of a selected exchange online organization Role "+"<b>$($SpecificRoleMembers.role[0].tostring())</b>"+", for more information on RBAC please check the following article: $article"
[PSCustomObject]$StartHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString "Please check the next section for more information!"
$null = $TheObjectToConvertToHTML.Add($StartHTML)
[string]$SectionTitle = "Management Role Assignment Users Table"
[string]$Description = "This section lists a table with the tenant users & their corresponding management RBAC roles."
[PSCustomObject]$RBACHTML = New-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $SpecificRoleMembers -TableType Table
$null = $TheObjectToConvertToHTML.Add($RBACHTML)
#region ResultReport
[string]$FilePath = $ExportPath + "\ManagementRoleAssignmentUsers.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Management Role Assignment Users" -ReportTitle "Management Role Assignment Users" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
#Question to ask enduser for opening the HTMl report
$OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile -like "*y*")
{
Write-Host "Opening report...." -ForegroundColor Cyan
Start-Process $FilePath
}
#endregion ResultReport
$SpecificRoleMembers|export-csv "$ExportPath\RoleMembers_$ts.csv" -NoTypeInformation 
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow
Read-Key
Start-O365TroubleshootersMenu