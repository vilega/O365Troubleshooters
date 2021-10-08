<# ------------------------------------------------------------------------------------------------------------------------------------------
Description: 
    Get retention policies and eDiscovery holds rules for SharePoint, OneDrive and ModernGroups (Microsoft 365 Groups)
    and exports the result for a CSV file if specified
    

Attention:
  - This script requires a Security & Compliance PS session active like described in:
    Without MFA - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/connect-to-scc-powershell
    With MFA - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell
  - This script DOES NOT check eDiscovery holds neither DLP policies
-------------------------------------------------------------------------------------------------------------------------------------------#>

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
Connect-IPPSSession -UserPrincipalName meganb@M365x768391.onmicrosoft.com
#>
#endregion 

# Fill in export path ---------------------------------
$exportPath = "C:\temp\SPOTenantHoldsReport.csv"
# -----------------------------------------------------

# Initialization variables
$workloads = @("SharePoint","OneDrive","ModernGroup") # SharePoint, OneDrive, ModernGroup (Microsoft 365 Groups)
$Report = @()

# Get the info for each workload specified, then parses each policy in them
Foreach ($workload in $workloads){
    
    #Loads the retention policies
    $Policies = (Get-RetentionCompliancePolicy -ExcludeTeamsPolicy -DistributionDetail | ? {$_.$($workload + "Location") -ne $Null})
    #Loads the holds from eDiscovery cases
    $Policies += Get-ComplianceCase | Foreach { Get-CaseHoldPolicy -Case $_.Identity -DistributionDetail | ? {$_.$($workload + "Location") -ne $Null} }
    
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
                If ($P.$($workload + "LocationException") -ne $Null) {
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
            ElseIf ($P.$($workload + "Location").Name -ne "All") {
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
If ($exportPath) { $Report | Export-Csv -NoTypeInformation $exportPath -Encoding UTF8 }
