Function new-AADSyncDDGRules {
Remove-Variable syncRule -Force

#region Create ADSync PROVISION rule
New-ADSyncRule  `
-Name 'In from AD - Dynamic Distribtution Group - Provision' `
-Description 'Dynamic Groups sync' `
-Direction 'Inbound' `
-Precedence 1 `
-PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
-PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
-SourceObjectType 'msExchDynamicDistributionList' `
-TargetObjectType 'person' `
-Connector '12ab56a5-9827-4480-b624-3e8af2fcce7d' `
-LinkType 'Provision' `
-SoftDeleteExpiryInterval 0 `
-OutVariable syncRule
#endregion Create ADSync PROVISION rule

#region Created Attribute Flow Mappings for the PROVISION rule
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'cloudFiltered' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsPresent([isCriticalSystemObject]) || ( (InStr([displayName], "(MSOL)") > 0) && (CBool([msExchHideFromAddressLists]))) || (Left([mailNickname], 4) = "CAS_" && (InStr([mailNickname], "}") > 0)) || CBool(InStr(DNComponent(CRef([dn]),1),"\\0ACNF:")>0), True, NULL)' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'mailEnabled' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(( (IsPresent([proxyAddresses]) = True) && (Contains([proxyAddresses], "SMTP:") > 0) && (InStr(Item([proxyAddresses], Contains([proxyAddresses], "SMTP:")), "@") > 0)) ||  (IsPresent([mail]) = True && (InStr([mail], "@") > 0)), True, False)' `
-OutVariable syncRule
#endregion Created Attribute Flow Mappings for the PROVISION rule

#region Create Scope Condition to match all Dynamic Distribtuion object type for PROVISION rule
New-Object  `
-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
-ArgumentList 'isCriticalSystemObject','True','NOTEQUAL' `
-OutVariable condition0


New-Object  `
-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
-ArgumentList 'msExchRecipientDisplayType','3','EQUAL' `
-OutVariable condition1


Add-ADSyncScopeConditionGroup  `
-SynchronizationRule $syncRule[0] `
-ScopeConditions @($condition0[0],$condition1[0]) `
-OutVariable syncRule

#endregion Create Scope Condition to match all Dynamic Distribtuion object type

#region Create Join Condition for PROVISION rule
New-Object  `
-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.JoinCondition' `
-ArgumentList 'mail','mail',$false `
-OutVariable condition0


Add-ADSyncJoinConditionGroup  `
-SynchronizationRule $syncRule[0] `
-JoinConditions @($condition0[0]) `
-OutVariable syncRule
#endregion Create Join Condition for PROVISION rule


Add-ADSyncRule -SynchronizationRule $syncRule[0]

Get-ADSyncRule -Identifier $syncRule[0].Identifier


Remove-Variable syncRule -Force




#region Create ADSync JOIN rule (Has All Attribute Transformations)
New-ADSyncRule  `
-Name 'In from AD - Dynamic Distribution Group - Join' `
-Description 'Dynamic Distribution List object with Exchange schema in Active Directory.' `
-Direction 'Inbound' `
-Precedence 3 `
-PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
-PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
-SourceObjectType 'msExchDynamicDistributionList' `
-TargetObjectType 'person' `
-Connector '12ab56a5-9827-4480-b624-3e8af2fcce7d' `
-LinkType 'Join' `
-SoftDeleteExpiryInterval 0 `
-OutVariable syncRule
#-Identifier '73838bb7-dd6c-42fe-b19f-b85e8d6808f6' `
#-ImmutableTag 'Microsoft.InfromADContactCommon.006' `
#endregion Create ADSync JOIN rule

#region Create Attribute Flow Mappings
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('dn') `
-Destination 'distinguishedName' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('assistant') `
-Destination 'assistant' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('authOrig') `
-Destination 'authOrig' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('cn') `
-Destination 'cn' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'company' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([company])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'department' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([department])' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'description' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsNullOrEmpty([description]),NULL,Left(Trim(Item([description],1)),448))' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'displayName' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsNullOrEmpty([displayName]),[cn],[displayName])' `
-OutVariable syncRule


# to verify what the attribute is doing
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('dLMemSubmitPerms') `
-Destination 'dLMemSubmitPerms' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


# to verify what the attribute is doing
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('dLMemRejectPerms') `
-Destination 'dLMemRejectPerms' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute1' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute1])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute2' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute2])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute3' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute3])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute4' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute4])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute5' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute5])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute6' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute6])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute7' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute7])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute8' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute8])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute9' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute9])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute10' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute10])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute11' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute11])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute12' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute12])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute13' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute13])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute14' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute14])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'extensionAttribute15' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([extensionAttribute15])' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'info' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Left(Trim([info]),448)' `
-OutVariable syncRule
#>


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'l' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([l])' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'legacyExchangeDN' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsPresent([legacyExchangeDN]), [legacyExchangeDN], NULL)' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'mail' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([mail])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'mailNickname' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsPresent([mailNickname]), [mailNickname], [cn])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msDS-HABSeniorityIndex') `
-Destination 'msDS-HABSeniorityIndex' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msDS-PhoneticDisplayName') `
-Destination 'msDS-PhoneticDisplayName' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchAssistantName') `
-Destination 'msExchAssistantName' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchBlockedSendersHash') `
-Destination 'msExchBlockedSendersHash' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchBypassModerationFromDLMembersLink') `
-Destination 'msExchBypassModerationFromDLMembersLink' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchBypassModerationLink') `
-Destination 'msExchBypassModerationLink' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchExtensionCustomAttribute1') `
-Destination 'msExchExtensionCustomAttribute1' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchExtensionCustomAttribute2') `
-Destination 'msExchExtensionCustomAttribute2' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchExtensionCustomAttribute3') `
-Destination 'msExchExtensionCustomAttribute3' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchExtensionCustomAttribute4') `
-Destination 'msExchExtensionCustomAttribute4' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchExtensionCustomAttribute5') `
-Destination 'msExchExtensionCustomAttribute5' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchHideFromAddressLists') `
-Destination 'msExchHideFromAddressLists' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchLitigationHoldDate') `
-Destination 'msExchLitigationHoldDate' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchLitigationHoldOwner') `
-Destination 'msExchLitigationHoldOwner' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchModeratedByLink') `
-Destination 'msExchModeratedByLink' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchModerationFlags') `
-Destination 'msExchModerationFlags' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchRecipientDisplayType') `
-Destination 'msExchRecipientDisplayType' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchRecipientTypeDetails') `
-Destination 'msExchRecipientTypeDetails' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchRequireAuthToSendTo') `
-Destination 'msExchRequireAuthToSendTo' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchRetentionComment') `
-Destination 'msExchRetentionComment' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchRetentionURL') `
-Destination 'msExchRetentionURL' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchSafeRecipientsHash') `
-Destination 'msExchSafeRecipientsHash' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchSafeSendersHash') `
-Destination 'msExchSafeSendersHash' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('msExchSenderHintTranslations') `
-Destination 'msExchSenderHintTranslations' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'proxyAddresses' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'RemoveDuplicates(Trim(ImportedValue("proxyAddresses")))' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'sourceAnchor' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'ConvertToBase64([objectGUID])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('objectGUID') `
-Destination 'sourceAnchorBinary' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('Contact') `
-Destination 'sourceObjectType' `
-FlowType 'Constant' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('targetAddress') `
-Destination 'targetAddress' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule
#>


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'telephoneNumber' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Trim([telephoneNumber])' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Source @('unauthOrig') `
-Destination 'unauthOrig' `
-FlowType 'Direct' `
-ValueMergeType 'Update' `
-OutVariable syncRule


<#
Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'url' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'IIF(IsNullOrEmpty([url]),NULL,Left(Trim(Item([url],1)),448))' `
-OutVariable syncRule
#>
#endregion Create Attribute Flow Mappings

#region Create Scope Condition to match all Dynamic Distribtuion object type
New-Object  `
-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
-ArgumentList 'msExchRecipientDisplayType','3','EQUAL' `
-OutVariable condition0


Add-ADSyncScopeConditionGroup  `
-SynchronizationRule $syncRule[0] `
-ScopeConditions @($condition0[0]) `
-OutVariable syncRule
#endregion Create Scope Condition to match all Dynamic Distribtuion object type


# Add the rule in the AAD Connect engine
Add-ADSyncRule -SynchronizationRule $syncRule[0]

Get-ADSyncRule  -Identifier $syncRule[0].Identifier

#Remove-Variable syncRule -Force


# Remove the rule in the AAD Connect engine
#Remove-ADSyncRule -SynchronizationRule $syncRule[0]
#Remove-Variable syncRule -Force

}

Function Get-AADSyncDDGRulesExists {

    $adConnectors = (Get-ADSyncConnector | ? {$_.Name -notlike "*onmicrosoft*"}).identifier.guid
    # ADSync rule with Dynamic Distribution Groups as source object
    $ADSyncDDGRules = Get-ADSyncRule | ? {$_.Direction -eq "Inbound" -and $_.Connector -eq $adConnectors -and $_.SourceObjectType -eq "msExchDynamicDistributionList" -and $_.TargetObjectType -eq "person"}

    # Verify if the AAD DDG Rule exists
    $AdSyncDDG_Join = $ADSyncDDGRules | ? {$_.LinkType -eq "Join"}
    foreach ($prop in ($AdSyncDDG_Join).AttributeFlowMappings) {
        if ($prop.Destination -eq "sourceObjectType" -and $prop.Source -contains "Contact" -and $prop.FlowType -eq "Constant") {
            $prop
        }
    }
    ($AdSyncDDG_Join).ScopeFilter.ScopeConditionList
}


Write-Host "This dianostic have to be executed on AAD Connect server to have access to AADSync PowerShell Module!" -ForegroundColor Yellow
Read-Key

Clear-Host
$Workloads = "AdSync"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\UnifiedAudit_$ts"
mkdir $ExportPath -Force |Out-Null

if (!(Get-AADSyncDDGRulesExists )) {
    new-AADSyncDDGRules
}

#TODO: write all steps in logs
#TODO: check what can be exported in HTML report 
#TODO: track if any error while creating the rules
#TODO: maybe would be better to check if the rules are implemented to show in Report

Read-Key
Start-O365TroubleshootersMenu
