#
# Module manifest for module 'PSGet_O365Troubleshooters'
#
# Generated by: vilega@microsoft.com
#
# Generated on: 6/1/2020
#

@{

# Script module or binary module file associated with this manifest.
RootModule = '.\O365Troubleshooters.psm1'

# Version number of this module.
ModuleVersion = '2.0.0.11'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '62b5fae4-8cf8-4390-91a4-a59fa53951bc'

# Author of this module
Author = 'vilega@microsoft.com'

# Company or vendor of this module
CompanyName = 'Victor Legat'

# Copyright statement for this module
Copyright = '(c) 2020 vilega@microsoft.com. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Office 365 Troubleshooters module has been designed to help Office 365 Administrators to do troubleshooting on Office 365 services. 
The module will be monthly updated with new features and improve existing ones.
The project is available at https://github.com/vilega/O365Troubleshooters
PowerShell 7 is not supported as some Office 365 connections modules are not yet compatible!'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = 'Start-O365Troubleshooters'

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
# VariablesToExport = @()

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
FileList = @(
    'ActionPlans\Export-ExoQuarantineMessages.ps1'
    'ActionPlans\Start-AllUsersWithAllRoles.ps1'
    'ActionPlans\Start-AzureADAuditSignInLogSearch.ps1'
    'ActionPlans\Start-CompromisedInvestigation.ps1'
    'ActionPlans\Start-DecodeSafeLinksURL.ps1'
    'ActionPlans\Start-ExchangeOnlineAuditSearch.ps1'
    'ActionPlans\Start-FindUserWithSpecificRbacRole.ps1'
    'ActionPlans\Start-MailboxDiagnosticLogs.ps1'
    'ActionPlans\Start-Office365Relay.ps1'
    'ActionPlans\Start-OfficeMessageEncryption.ps1'
    'ActionPlans\Start-RbacTools.ps1'
    'ActionPlans\Start-UnifiedAuditLogSearch.ps1'
)

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @("O365","Office365","Exchange","EXO","ExchangeOnline","Compliance","Security","Identity","Audit","OME","AIP","OfficeMessageEncryption","Azure","Protection","AzureInformationProtection","UnifiedLabeling","Diagnostic","Actionplan","Report","Tool")

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/VictorLegat/O365Troubleshooters'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = '
        2.0.0.11 - Fixed export of AzureAD Audit SignIn Logs
        2.0.0.10 - Fixed issue with Decode SafeLinks URL
        2.0.0.7 - Added Azure Sign in logs
        2.0.0.6 - Module deployed with main functions and the following action plans: OME, SMTP Relay, EXO Audit, Unified Audit, Find All Users with Specific RBAC, Export all users with RBAC, Export Mailbox Diagnostics logs, Decode ATP SafeLinks, Export Quarantine
        '

        # External dependent modules of this module
        # ExternalModuleDependencies = ''

    } # End of PSData hashtable
    
 } # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''
}

