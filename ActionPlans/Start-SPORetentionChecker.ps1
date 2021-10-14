<#
    .SYNOPSIS
    Provide a report for retention policies and their distribution status along with recommendation for each case.

    .DESCRIPTION
    Provide a report for retention policies and eDiscovery holds rules affecting SharePoint, OneDrive and Microsoft 365 Groups.
    Check distribution status for the associated policies.
    Provide recommended actions and reference given the main scenarios.

    .EXAMPLE
    Use a global admin when prompted for a account.
    
    .LINK
    Online documentation: https://answers.microsoft.com/

#>

#region To-Do --------------------------------------------------------
<#
ok 1 - Enumerate common issues this script may spot
ok 2 - Add output to highlight the issues and guide fixes
ok 3 - Integrate with main menu
ok 4 - Integrate with S&C connection function --> Pending function fix worked by Victor
ok 5 - Code the issues checks scoped for hackathon
   6 - Create wiki for the Action Plan
   7 - Create the help info using the template
   x - Add eDiscovery case name and ID to CaseHold items
   x - Automate adding exception for a site?
   x - Automate adding exception for a M365 group?

1 - Enumerate common issues this script may spot
    ok  a. Sharepoint holds over ODB sites
    ok  b. M365 group holds over SPO sites
    ok  c. Distribution issues checking status for Mode, DistributionStatus, DistributionResults
        d. Show policies protecting a site --> Create issue for later (not in hackathon sprint) 

2 - Add output to highlight the issues and guide fixes
    ok  a. Add references about how to setup the exclusions
            https://docs.microsoft.com/en-us/sharepoint/troubleshoot/administration/exclude-sites-from-retention-policy
    ok  b. Explain about the grace-period for disabled policies
            https://docs.microsoft.com/en-us/microsoft-365/compliance/retention?view=o365-worldwide#releasing-a-policy-for-retention
    ok  c. Add reference to the article to remove inconsistent policies
            https://docs.microsoft.com/en-us/sharepoint/troubleshoot/sites/compliance-policy-blocking-site-deletion
    ok  d. Add reference about retry distribution
            https://docs.microsoft.com/en-us/powershell/module/exchange/set-retentioncompliancepolicy
    ok  e. Retention labels
            https://docs.microsoft.com/en-us/microsoft-365/compliance/retention?view=o365-worldwide#retention-policies-and-retention-labels

5 - Code the issues checks scoped for hackathon
    ok  a. Sharepoint holds over ODB sites
    ok  b. M365 group holds over SPO sites
    ok  c. Distribution issues checking status for Mode, DistributionStatus, DistributionResults

6 - Create wiki for the Action Plan
    Word document + share with Victor
    https://answers.microsoft.com/en-us/msoffice/forum/all/how-to-diagnose-invalid-public-folder-dumpster/55d5fdfc-9309-4764-9b2e-5466bedc5227
    https://answers.microsoft.com/en-us/msoffice/forum/all/validating-distribution-group-eligibility-for/dd3a2271-cb97-4579-8935-19409dab1dc2
    https://answers.microsoft.com/en-us/msoffice/forum/all/creating-aad-connect-rules-to-synchronize-on/41444825-f62f-4f1a-a449-152806319568

Questions:
    Closed   1 - Field order in report --> Victor will either create a switch or let it respect the object order
    Closed   2 - Export CSV should be optional? --> Export when not empty
    Closed   3 - Make the report more inclusive including signs rather than colors only --> We'll create an 'issue' to build this after the hackathon 
    Closed   4 - PS session crashing --> Due to the module design, it won't happen in prod experience

Issues:
    1 - Script does not identify Enabled policies which have the workload disabled (hold will remain during grace-period)
    2 - Show policies protecting a site --> Create issue for later (not in hackathon sprint) 
#>
#endregion ------------------------------------------------------------

<# DEV Setup ------------------------------------------------------
import-module C:\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
#Start-O365TroubleshootersMenu
Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com -Prefix cc
-------------------------------------------------------------------   #>

#Connect-O365PS "SCC"
Connect-IPPSSession -Prefix cc
#Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com -Prefix cc


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

#region Get data
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
                  PolicyName          = $P.Name
                  SiteName            = "All $workload Sites"
                  Address             = "All $workload Sites"
                  Workload            = $workload
                  Type                = $P.Type
                  Guid                = $P.Guid
                  Mode                = $P.Mode
                  DistributionStatus  = $P.DistributionStatus
                  DistributionResults = $P.DistributionResults
                  Enabled             = $P.Enabled
                  }
                $Report += $ReportLine 

                # Check if the policy have exceptions
                If ($P.$($workload + "LocationException").count -gt 0) {
                    $Locations = ($P | Select-Object -ExpandProperty $($workload + "LocationException"))
                        ForEach ($L in $Locations) {
                            $Exception = "[EXCLUDE] " + $L.DisplayName
                            $ReportLine = [PSCustomObject]@{
                                            PolicyName          = $P.Name
                                            SiteName            = $Exception
                                            Address             = $L.Name
                                            Workload            = $workload
                                            Type                = $P.Type
                                            Guid                = $P.Guid
                                            Mode                = $P.Mode
                                            DistributionStatus  = $P.DistributionStatus
                                            DistributionResults = $P.DistributionResults
                                            Enabled             = $P.Enabled
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
                                    PolicyName          = $P.Name
                                    SiteName            = $L.DisplayName
                                    Address             = $L.Name
                                    Workload            = $workload
                                    Type                = $P.Type
                                    Guid                = $P.Guid
                                    Mode                = $P.Mode
                                    DistributionStatus  = $P.DistributionStatus
                                    DistributionResults = $P.DistributionResults
                                    Enabled             = $P.Enabled
                                    }
                    $Report += $ReportLine  
                }                    
            }
    }
}
#endregion Get data

#region Treat data
# Shape Report Source Objects
$HoldsReport = $Report | Select-Object PolicyName, SiteName, Address, Workload, Type, Guid
$PoliciesReport = $Report | 
                    Sort-Object Guid -Unique | 
                        Select-Object PolicyName, Type, Enabled, Guid, Mode, DistributionStatus, DistributionResults
# Removes 'DistributionResults' to avoid polluting the report
$PoliciesLeanReport = $PoliciesReport | Select-Object PolicyName, Type, Enabled, Guid, Mode, DistributionStatus

# Filter SharePoint Holds - SharePoint + M365 groups
$SPOReport = New-Object -TypeName "System.Collections.ArrayList"
$SPOReport = $HoldsReport | Where-Object {$_.Workload -in @("SharePoint","ModernGroup")} 
$SPOReport = $SPOReport | Sort-Object -Descending "SiteName" | Sort-Object "PolicyName"

# Filter OneDrive Holds - SharePoint and OneDrive except policies expliciting only SharePoint Sites
$ODBReport = New-Object -TypeName "System.Collections.ArrayList"
$ODBReport = $HoldsReport | Where-Object {$_.Workload -eq "SharePoint" -and $_.Address -notmatch "sharepoint.com"}
$ODBReport += $HoldsReport | Where-Object {$_.Workload -eq "SharePoint" -and $_.Address -match "-my.sharepoint.com"}
$ODBReport += $HoldsReport | Where-Object {$_.Workload -eq "OneDrive"}
$ODBReport = $ODBReport | Sort-Object -Descending "SiteName" | Sort-Object "PolicyName"

# Filter Healthy Policies
$HealthyPolicies = New-Object -TypeName "System.Collections.ArrayList"
$HealthyPolicies += $PoliciesLeanReport | 
                        Where-Object {
                            $_.DistributionStatus -eq "Success" -and
                            $_.DistributionResult.count -eq 0 -and
                            $_.Mode -eq "Enforce"
                        }
                        #$Policies[0] | FL

# Filter Distribution issues (Mode not: 'Enforce'; Status not: 'Success';  Results not: empty)
$DistributionIssues = New-Object -TypeName "System.Collections.ArrayList"
$DistributionIssues += $PoliciesReport | 
                        Where-Object {
                            $_.DistributionStatus -ne "Success" -or
                            $_.DistributionResult.count -gt 0 -or 
                            $_.Mode -ne "Enforce"
                        }

#endregion Treat data

#region Present data
$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"

# Adds general guidance about the report
[string]$SectionTitle = "Report Guidance"
[string]$Description = 'Consider these principles to interprete the information in this report:
                        <ul style="margin:0 0 0 10px">
                            <li>A Site/OneDrive explicitly included is protected.</li>
                            <li>A Site/OneDrive included by an "All [workload] Sites"  policy is protected if not explicitly excluded by the same policy.</li>
                            <li>Inclusion policies precede exclusion policies.</li>
                        </ul>'
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString ""

$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds the session for healthy policies  in case there is any
If ($HealthyPolicies.Count -gt 0){
    [string]$SectionTitle = "Healthy Policies"
    [string]$Description = 'These policies were checked for common distribution problems and no issue were found.<br>
                            <b>Note:</b> Disabled policies may still enforce holds due to the grace-period.
                            (Learn more about <a href="https://docs.microsoft.com/en-us/microsoft-365/compliance/retention?view=o365-worldwide#releasing-a-policy-for-retention" target="_blank">Grace-Period</a>)'
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $HealthyPolicies -TableType Table
    $null = $TheObjectToConvertToHTML.Add($SectionHtml)
}

# Adds the session for distribution issues in case there is any
If ($DistributionIssues.Count -gt 0){
    # Adds the session for distribution status
    [string]$SectionTitle = "Policies Distribution Issues"
    [string]$Description = 'The policies below are in an inconsistent status, review and consider fixing the issues pointed out in <i>DISTRIBUTIONRESULTS</i> column. Then, retry the distribution.<br>
                            Use the following command to retry distribution for a given policy:<br>
                            <div style="margin-left:20px">
                            <i>
                                Set-RetentionCompliancePolicy -Identity "Hold Everything" -RetryDistribution    <br>
                                Set-CaseHoldPolicy            -Identity "20ab020d-1c7e-464e-85f7-ef720c41825d" -RetryDistribution
                            </i>
                            <div>'
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $DistributionIssues -TableType Table
    $null = $TheObjectToConvertToHTML.Add($SectionHtml)
}

# Adds the session for policies affecting SharePoint
[string]$SectionTitle = "SharePoint Holds"
If ($SPOReport.Count -gt 0){
    [string]$Description = "These are the holds which may be preventing SharePoint files and sites to be deleted.<br>
                            Sites connected to a M365 group are impacted by ModernGroup holds."
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $SPOReport -TableType Table
} Else {
    [string]$Description = 'No retention policy or eDiscovery case hold were found that could be preventing SharePoint files and sites to be deleted.<br>
                            In case you still face issues consider checking retention labels and disabled policies in grace-period.
                            (Learn more about <a href="https://docs.microsoft.com/en-us/microsoft-365/compliance/retention?view=o365-worldwide#retention-policies-and-retention-labels" target="_blank">Retention Labels</a>)'
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString ""   
}
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds the session for policies affecting OneDrive
[string]$SectionTitle = "OneDrive Holds"
If ($ODBReport.Count -gt 0){   
    [string]$Description = "These are the holds which may be preventing OneDrive files and sites to be deleted."
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $ODBReport -TableType Table
} Else {
    [string]$Description = 'No retention policy or eDiscovery case hold were found that could be preventing OneDrive files and sites to be deleted.<br>
                            In case you still face issues consider checking retention labels and disabled policies in grace-period.
                            (Learn more about <a href="https://docs.microsoft.com/en-us/microsoft-365/compliance/retention?view=o365-worldwide#retention-policies-and-retention-labels" target="_blank">Retention Labels</a>)'
    [PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString ""
}
$null = $TheObjectToConvertToHTML.Add($SectionHtml)

# Adds recomendations to the report
[string]$SectionTitle = "Recommendations"
[string]$Description = 'To delete sites or their files, you must add an exclusion in all policies affecting the site.
                        Check how to <a href="https://docs.microsoft.com/en-us/sharepoint/troubleshoot/administration/exclude-sites-from-retention-policy" target="_blank">Exclude Sites from Retention Policy</a>. <br>
                        <b>Note:</b> For sites connected to Microsoft 365 group you must also exclude them from the M365 group workload associated.'
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDataString ""
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
#endregion Present data
    
# Print location where the data was exported
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Read-Key

Start-O365TroubleshootersMenu