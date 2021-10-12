<# ------------------------------------------------------------------------------------------------------------------------------------------
Description: 
    Get retention policies and eDiscovery holds rules for SharePoint, OneDrive and ModernGroups (Microsoft 365 Groups)
    and exports the result for a CSV file if specified

Issues:
- SPO site bounded to a M365 non-excluded from from ModernGroups policies (Accept site URL parameter?)
- Find a eDiscovery hold applied for a SPO/ODB site (Accept site URL parameter?)

Attention:
  - This script requires a Security & Compliance PS session active like described in:
    Without MFA - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/connect-to-scc-powershell
    With MFA - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell
  - This script DOES NOT check eDiscovery holds neither DLP policies
-------------------------------------------------------------------------------------------------------------------------------------------#>

#region To-Do --------------------------------------------------------
<#

1 - Enumerate common issues this script may spot
2 - Add output to highlight the issues and guide fixes
3 - Integrate with main menu
4 - Integrate with S&C connection function
5 - Create wiki for the Action Plan
x - Automate adding exception for a site?
x - Automate adding exception for a M365 group?

#>
#endregion ------------------------------------------------------------


#region Quick reference on connectiong steps
<#
## Without MFA
$UserCredential = Get-Credential -UserName "meganb@M365x768391.onmicrosoft.com" -Message "Enter password:"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

# Always remove the session otherwise it may lock you out and you'll need to wait previous sessions expire
Remove-PSSession $Session


## With MFA
# Exchange Admin Center > Hybrid > Configure, install...
Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com
#>
#endregion 


<# DEv Setup
import-module C:\GitHub\O365Troubleshooters\O365Troubleshooters.psm1 -Force
Set-GlobalVariables
Start-O365TroubleshootersMenu
#>

#Connect-O365PS "SCC"
#Connect-IPPSSession -UserPrincipalName roan@roanmarques.onmicrosoft.com -Prefix cc

# Fill in export path ---------------------------------


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
$Report = @()
$Policies = @()

# Get the info for each workload specified, then parses each policy in them
Foreach ($workload in $workloads){
    
    #Loads the retention policies
    $Policies += (Get-ccRetentionCompliancePolicy -ExcludeTeamsPolicy -DistributionDetail | Where-Object {$_.$($workload + "Location") -ne $Null})
    #Loads the holds from eDiscovery cases
    $Policies += Get-ccComplianceCase | ForEach-Object { Get-ccCaseHoldPolicy -Case $_.Identity -DistributionDetail |  Where-Object {$_.$($workload + "Location") -ne $Null} }
    
    # Parses each policy info
    ForEach ($P in $Policies) {
            
            # Treat the policies where the scope is the whole workloads, and following check for exceptions
            If ($P.$($workload + "Location").Name -eq "All") {
                $ReportLine = [PSCustomObject]@{
                  PolicyName = $P.Name
                  SiteName   = "All $workload Sites"
                  Address    = "All $workload Sites"
                  Workload   = "$workload"
                  Type       = $P.Type
                  Guid       = $P.Guid
                  }
                $Report += $ReportLine }

                # Check if the policy have exceptions
                If ($P.$($workload + "LocationException").count -gt 0) {
                    $Locations = ($P | Select -ExpandProperty $($workload + "LocationException"))
                        ForEach ($L in $Locations) {
                            $Exception = "*Exclude* " + $L.DisplayName
                            $ReportLine = [PSCustomObject]@{
                                            PolicyName = $P.Name
                                            SiteName   = $Exception
                                            Address    = $L.Name
                                            Workload   = "$workload"
                                            Type       = $P.Type
                                            Guid       = $P.Guid
                                        }
                            $Report += $ReportLine
                        }
            }

            # Tread the policies where the scope is restrict to specific locations in the given workload
            If ($P.$($workload + "Location").Name -ne "All") {
                $Locations = ($P | Select -ExpandProperty $($workload + "Location"))
                ForEach ($L in $Locations) {
                    $ReportLine = [PSCustomObject]@{
                                    PolicyName = $P.Name
                                    SiteName   = $L.DisplayName
                                    Address    = $L.Name
                                    Workload   = "$workload"
                                    Type       = $P.Type
                                    Guid       = $P.Guid
                                    }
                    $Report += $ReportLine  
                }                    
            }
    }
}

# Shows the reports in a gridview window
$Report | Out-GridView -Title "SPO Tenant holds Report"

# If specified, exports the report for an CSV file  
If ($exportPath) { $Report | Export-Csv -NoTypeInformation $ExportPath\SPOTenantHoldsReport.csv -Encoding UTF8 }

$TheObjectToConvertToHTML = New-Object -TypeName "System.Collections.ArrayList"
[string]$SectionTitle = "Policies"
[string]$Description = "If there is a hold, ...."
#[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDataString "Bla bla"
[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Report -TableType Table
#[PSCustomObject]$SectionHtml = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "CustomObject" -EffectiveDataArrayList $Report -TableType List
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
#endregion ResultReport
    
# Print location where the data was exported
Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 
Read-Key

Start-O365TroubleshootersMenu

$wor