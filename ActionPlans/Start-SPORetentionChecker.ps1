<# ------------------------------------------------------------------------------------------------------------------------------------------
Description: 
    Get retention policies and eDiscovery holds rules for SharePoint, OneDrive and ModernGroups (Microsoft 365 Groups)

-------------------------------------------------------------------------------------------------------------------------------------------#>

#region To-Do --------------------------------------------------------
<#

ok 1 - Enumerate common issues this script may spot
2 - Add output to highlight the issues and guide fixes
3 - Integrate with main menu
4 - Integrate with S&C connection function
5 - Create wiki for the Action Plan
    word document
x - Automate adding exception for a site?
x - Automate adding exception for a M365 group?

1 - Enumerate common issues this script may spot
  ok  a. Sharepoint holds over ODB sites
  ok  b. M365 group holds over SPO sites
    c. Distribution issues checking status for Mode, DistributionStatus, DistributionResults
    d. Show policies protecting a site

2 - Add output to highlight the issues and guide fixes
    a. Add references about how to setup the exclusions


Questions:
    1 - Field order in report
    2 - Export CSV should be optional?
    3 - Make the report more inclusive
    4 - PS session crashing --> Due to the module design
    
#>
#endregion ------------------------------------------------------------

<# DEV Setup ------------------------------------------------------
import-module C:\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
#Start-O365TroubleshootersMenu
Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com -Prefix cc
-------------------------------------------------------------------   #>

#Connect-O365PS "SCC"
#Connect-IPPSSession -Prefix cc
Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com -Prefix cc


# Create the Export Folder
$ts = get-date -Format yyyyMMdd_HHmmss

# Create export folder
try {
    $ExportPath = "$global:WSPath\SPOTenantHoldsReport_$ts"
    mkdir $ExportPath -Force | out-null
    Write-Log -function "Start-SPORetentionChecker" -step  "Create ExportPath" -Description "Success"
}
catch {
    Write-Log -function "Start-SPORetentionChecker" -step  "Create ExportPath" -Description "Couldn't create folder $global:WSPath\SPOTenantHoldsReport_$ts. Error: $($_.Exception.Message)"
    Write-Host "Couldn't create folder $global:WSPath\SPOTenantHoldsReport_$ts"
    Read-Key
    Start-O365TroubleshootersMenu
}



# -----------------------------------------------------

# Initialization variables
$workloads = @("SharePoint","OneDrive","ModernGroup") # SharePoint, OneDrive, ModernGroup (Microsoft 365 Groups)
$Report = New-Object -TypeName "System.Collections.ArrayList"


# Get the info for each workload specified, then parses each policy in them
Foreach ($workload in $workloads){
    
    #Loads the retention policies
    $Policies = New-Object -TypeName "System.Collections.ArrayList" #$Policies must be reset for each workload and it must be and array
    $Policies += (Get-ccRetentionCompliancePolicy -ExcludeTeamsPolicy -DistributionDetail | Where-Object {$_.$($workload + "Location") -ne $Null})
    #Loads the holds from eDiscovery cases
    $Policies += Get-ccComplianceCase | 
        ForEach-Object { Get-ccCaseHoldPolicy -Case $_.Identity -DistributionDetail |
            Where-Object {$_.$($workload + "Location") -ne $Null} }
    
    # Parses each policy info
    ForEach ($P in $Policies) {
            
            # Treat the policies where the scope is the whole workloads, and following check for exceptions
            If ($P.$($workload + "Location").Name -eq "All") {
                $ReportLine = [PSCustomObject]@{
                  PolicyName = $P.Name
                  SiteName   = "All $workload Sites"
                  Address    = "All $workload Sites"
                  Workload   = "$workload"
                  #Type       = $P.Type #currently useless
                  Guid       = $P.Guid
                  }
                $Report += $ReportLine 

                # Check if the policy have exceptions
                If ($P.$($workload + "LocationException").count -gt 0) {
                    $Locations = ($P | Select-Object -ExpandProperty $($workload + "LocationException"))
                        ForEach ($L in $Locations) {
                            $Exception = "*Exclude* " + $L.DisplayName
                            $ReportLine = [PSCustomObject]@{
                                            PolicyName = $P.Name
                                            SiteName   = $Exception
                                            Address    = $L.Name
                                            Workload   = "$workload"
                                            #Type       = $P.Type #currently useless
                                            Guid       = $P.Guid
                                        }
                            $Report += $ReportLine
                        }
                }
            }

            # Treat the policies where the scope is restrict to specific locations in the given workload
            If ($P.$($workload + "Location").Name -ne "All") {
                $Locations = ($P | Select-Object -ExpandProperty $($workload + "Location"))
                ForEach ($L in $Locations) {
                    $ReportLine = [PSCustomObject]@{
                                    PolicyName = $P.Name
                                    SiteName   = $L.DisplayName
                                    Address    = $L.Name
                                    Workload   = "$workload"
                                    #Type       = $P.Type #currently useless
                                    Guid       = $P.Guid
                                    }
                    $Report += $ReportLine  
                }                    
            }
    }
}

# Filter SharePoint Holds - SharePoint + M365 groups
$SPOReport = $Report | Where-Object {$_.Workload -in @("SharePoint","ModernGroup")}
$SPOReport = $SPOReport | Sort-Object -Descending "SiteName" | Sort-Object "PolicyName"

# Filter OneDrive Holds - SharePoint and OneDrive except policies expliciting only SharePoint Sites
$ODBReport = New-Object -TypeName "System.Collections.ArrayList"
$ODBReport = $Report | Where-Object {$_.Workload -eq "SharePoint" -and $_.Address -notmatch "sharepoint.com"}
$ODBReport += $Report | Where-Object {$_.Workload -eq "SharePoint" -and $_.Address -match "-my.sharepoint.com"}
$ODBReport += $Report | Where-Object {$_.Workload -eq "OneDrive"}
$ODBReport = $ODBReport | Sort-Object -Descending "SiteName" | Sort-Object "PolicyName"


# If specified, exports the report for an CSV file  
$Report | Export-Csv -NoTypeInformation $ExportPath\SPOTenantHoldsReport.csv -Encoding UTF8

#region Prepare HTML Report
$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"

# HTML sample: Start-DlToO365GroupUpgradeChecks

# Adds general guidance about the report
[string]$SectionTitle = "Report Guidance"
[string]$Description = "A Site/OneDrive explicitly included is protected. `nA Site/OneDrive included by an 'All' workload policy is protected if not explicitly excluded by the same policy. `nInclusion policies precede exclusion policies."
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -DataType "String" -EffectiveDataString $Description
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds the session for policies affecting SharePoint
#TODO: check if there is content
[string]$SectionTitle = "SharePoint Holds"
[string]$Description = "These are the holds which may be preventing SharePoint files and sites to be deleted. Sites connected to a M365 group are impacted by ModernGroup holds."
#[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "Bla bla"
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $SPOReport -TableType Table
#[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Report -TableType List
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds the session for policies affecting OneDrive
[string]$SectionTitle = "OneDrive Holds"
[string]$Description = "These are the holds which may be preventing OneDrive files and sites to be deleted."
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ODBReport -TableType Table
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds recomendations to the report
[string]$SectionTitle = "Recommendations"
[string]$Description = "In case the Site/OneDrive you can't delete or remove files is not exempted in all the policies in which it's subjected to their scope, you'll need to add those exceptions to remove the retention protection."
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -DataType "String" -EffectiveDataString $Description
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

#Build HTML report out of the previous HTML sections
[string]$FilePath = $ExportPath + "\SPOTenantHoldsReport.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "SPO Tenant holds Report" -ReportTitle "SPO Tenant holds Report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML

#Ask end-user for opening the HTMl report
$OpenHTMLfile = Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile.ToLower() -like "*y*") {
    Write-Host "Opening report...." -ForegroundColor Cyan
    Start-Process $FilePath
}
#endregion Prepare HTML Report
    
# Print location where the data was exported
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Read-Key

Start-O365TroubleshootersMenu