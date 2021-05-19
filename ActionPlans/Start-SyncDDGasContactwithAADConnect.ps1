Function new-AADSyncDDGRules {

    #region Create ADSync PROVISION rule
    $ruleName = 'Custom In from AD - Dynamic Distribtution Group - Provision'
    if (!(Get-ADSyncRule -Identity $ruleName -ErrorAction SilentlyContinue)) {

    

        $slot = Get-ADSyncRuleFreeSlot -ruleName $ruleName 
        if ($slot.count -eq 0) {
            Write-Host "You select CANCEL so the rule `"$ruleName`" won't be created"
            #TODO: write this on the report
        }
        else {
    

            New-ADSyncRule  `
                -Name $ruleName  `
                -Description 'Dynamic Distribution Groups provision rule' `
                -Direction 'Inbound' `
                -Precedence $($slot.precedence) `
                -PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
                -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
                -SourceObjectType 'msExchDynamicDistributionList' `
                -TargetObjectType 'person' `
                -Connector '12ab56a5-9827-4480-b624-3e8af2fcce7d' `
                -LinkType 'Provision' `
                -SoftDeleteExpiryInterval 0 `
                -OutVariable syncRule | Out-Null
            #endregion Create ADSync PROVISION rule

            #region Created Attribute Flow Mappings for the PROVISION rule
            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'cloudFiltered' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(IsPresent([isCriticalSystemObject]) || ( (InStr([displayName], "(MSOL)") > 0) && (CBool([msExchHideFromAddressLists]))) || (Left([mailNickname], 4) = "CAS_" && (InStr([mailNickname], "}") > 0)) || CBool(InStr(DNComponent(CRef([dn]),1),"\\0ACNF:")>0), True, NULL)' `
                -OutVariable syncRule | Out-Null

            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'mailEnabled' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(( (IsPresent([proxyAddresses]) = True) && (Contains([proxyAddresses], "SMTP:") > 0) && (InStr(Item([proxyAddresses], Contains([proxyAddresses], "SMTP:")), "@") > 0)) ||  (IsPresent([mail]) = True && (InStr([mail], "@") > 0)), True, False)' `
                -OutVariable syncRule | Out-Null
            #endregion Created Attribute Flow Mappings for the PROVISION rule

            #region Create Scope Condition to match all Dynamic Distribtuion object type for PROVISION rule
            New-Object  `
                -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
                -ArgumentList 'isCriticalSystemObject', 'True', 'NOTEQUAL' `
                -OutVariable condition0 | Out-Null

            New-Object  `
                -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
                -ArgumentList 'msExchRecipientDisplayType', '3', 'EQUAL' `
                -OutVariable condition1  | Out-Null

            Add-ADSyncScopeConditionGroup  `
                -SynchronizationRule $syncRule[0] `
                -ScopeConditions @($condition0[0], $condition1[0]) `
                -OutVariable syncRule  | Out-Null

            #endregion Create Scope Condition to match all Dynamic Distribtuion object type

            #region Create join condition for PROVISION rule
            New-Object  `
                -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.JoinCondition' `
                -ArgumentList 'mail', 'mail', $false `
                -OutVariable condition0  | Out-Null

            Add-ADSyncJoinConditionGroup  `
                -SynchronizationRule $syncRule[0] `
                -JoinConditions @($condition0[0]) `
                -OutVariable syncRule  | Out-Null
            #endregion Create Join Condition for PROVISION rule

            #region Create the PROVISION rule
            Add-ADSyncRule -SynchronizationRule $syncRule[0] | Out-Null
            Remove-Variable syncRule -Force
        }
    }
    else {
        {
            Write-Host "Rule `"$ruleName`" already exist. If you want to re-create it, you can delete it from AAD Connect and re-run the script to re-created with the latest version"
            Read-Key

            Write-Host "You choosed not to go further for implementing the AADConnect rules to sync on-premises DDG to AAD/EXO Contacts" -ForegroundColor Red
            $CurrentProperty = "New-AADSyncDDGRules"
            $CurrentDescription = "rule `"$ruleName`" was already created"
            write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
            Write-Host "The script will return to main menu"
            read-Key    
            Start-O365TroubleshootersMenu
        }
    }
    #endregion Create the PROVISION rule


    #region Create ADSync JOIN rule (Has All Attribute Transformations)
    $ruleName = 'Custom In from AD - Dynamic Distribution Group - Join'
    if (!(Get-ADSyncRule -Identity $ruleName -ErrorAction SilentlyContinue)) {

        $slot = Get-ADSyncRuleFreeSlot -ruleName $ruleName 
        if ($slot.count -eq 0) {
            Write-Host "You select CANCEL so the rule `"$ruleName`" won't be created"
            #TODO: write this on the report
        }
        else {
            New-ADSyncRule  `
                -Name $ruleName  `
                -Description 'Dynamic Distribution Group Join - Transformations rule' `
                -Direction 'Inbound' `
                -Precedence $($slot.precedence) `
                -PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
                -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
                -SourceObjectType 'msExchDynamicDistributionList' `
                -TargetObjectType 'person' `
                -Connector '12ab56a5-9827-4480-b624-3e8af2fcce7d' `
                -LinkType 'Join' `
                -SoftDeleteExpiryInterval 0 `
                -OutVariable syncRule  | Out-Null
            #endregion Create ADSync JOIN rule

            #region Create Attribute Flow Mappings
            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('dn') `
                -Destination 'distinguishedName' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('authOrig') `
                -Destination 'authOrig' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('cn') `
                -Destination 'cn' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'description' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(IsNullOrEmpty([description]),NULL,Left(Trim(Item([description],1)),448))' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'displayName' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(IsNullOrEmpty([displayName]),[cn],[displayName])' `
                -OutVariable syncRule  | Out-Null


            # to verify what the attribute is doing
            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('dLMemSubmitPerms') `
                -Destination 'dLMemSubmitPerms' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            # to verify what the attribute is doing
            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('dLMemRejectPerms') `
                -Destination 'dLMemRejectPerms' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute1' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute1])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute2' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute2])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute3' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute3])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute4' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute4])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute5' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute5])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute6' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute6])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute7' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute7])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute8' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute8])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute9' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute9])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute10' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute10])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute11' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute11])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute12' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute12])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute13' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute13])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute14' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute14])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'extensionAttribute15' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([extensionAttribute15])' `
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'legacyExchangeDN' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(IsPresent([legacyExchangeDN]), [legacyExchangeDN], NULL)' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'mail' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'Trim([mail])' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'mailNickname' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'IIF(IsPresent([mailNickname]), [mailNickname], [cn])' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msDS-HABSeniorityIndex') `
                -Destination 'msDS-HABSeniorityIndex' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msDS-PhoneticDisplayName') `
                -Destination 'msDS-PhoneticDisplayName' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchExtensionCustomAttribute1') `
                -Destination 'msExchExtensionCustomAttribute1' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchExtensionCustomAttribute2') `
                -Destination 'msExchExtensionCustomAttribute2' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchExtensionCustomAttribute3') `
                -Destination 'msExchExtensionCustomAttribute3' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchExtensionCustomAttribute4') `
                -Destination 'msExchExtensionCustomAttribute4' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchExtensionCustomAttribute5') `
                -Destination 'msExchExtensionCustomAttribute5' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchHideFromAddressLists') `
                -Destination 'msExchHideFromAddressLists' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchRecipientDisplayType') `
                -Destination 'msExchRecipientDisplayType' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchRecipientTypeDetails') `
                -Destination 'msExchRecipientTypeDetails' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchRequireAuthToSendTo') `
                -Destination 'msExchRequireAuthToSendTo' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('msExchSenderHintTranslations') `
                -Destination 'msExchSenderHintTranslations' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'proxyAddresses' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'RemoveDuplicates(Trim(ImportedValue("proxyAddresses")))' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Destination 'sourceAnchor' `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression 'ConvertToBase64([objectGUID])' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('objectGUID') `
                -Destination 'sourceAnchorBinary' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('Contact') `
                -Destination 'sourceObjectType' `
                -FlowType 'Constant' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule | Out-Null


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


            #endregion Create Attribute Flow Mappings

            #region Create Scope Condition to match all Dynamic Distribtuion object type
            New-Object  `
                -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
                -ArgumentList 'msExchRecipientDisplayType', '3', 'EQUAL' `
                -OutVariable condition0

            Add-ADSyncScopeConditionGroup  `
                -SynchronizationRule $syncRule[0] `
                -ScopeConditions @($condition0[0]) `
                -OutVariable syncRule
            #endregion Create Scope Condition to match all Dynamic Distribtuion object type


            # Add the rule in the AAD Connect engine
            Add-ADSyncRule -SynchronizationRule $syncRule[0] | Out-Null
            Remove-Variable syncRule -Force
        }
    }
    else {
            Write-Host "Rule `"$ruleName`" already exist. If you want to re-create it, you can delete it from AAD Connect and re-run the script to re-created with the latest version"
            Read-Key
    
            Write-Host "You choosed not to go further for implementing the AADConnect rules to sync on-premises DDG to AAD/EXO Contacts" -ForegroundColor Red
            $CurrentProperty = "New-AADSyncDDGRules"
            $CurrentDescription = "rule `"$ruleName`" was already created"
            write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
            Write-Host "The script will return to main menu"
            read-Key    
            Start-O365TroubleshootersMenu
    }
}

Function Get-AADSyncDDGRulesExists {

    $adConnectors = (Get-ADSyncConnector | Where-Object { $_.Name -notlike "*onmicrosoft*" }).identifier.guid
    # ADSync rule with Dynamic Distribution Groups as source object
    $ADSyncDDGRules = Get-ADSyncRule | Where-Object { $_.Direction -eq "Inbound" -and $_.Connector -eq $adConnectors -and $_.SourceObjectType -eq "msExchDynamicDistributionList" -and $_.TargetObjectType -eq "person" }

    # Verify if the AAD DDG Rule exists
    $AdSyncDDG_Join = $ADSyncDDGRules | Where-Object { $_.LinkType -eq "Join" }
    foreach ($prop in ($AdSyncDDG_Join).AttributeFlowMappings) {
        if ($prop.Destination -eq "sourceObjectType" -and $prop.Source -contains "Contact" -and $prop.FlowType -eq "Constant") {
            $prop
        }
    }
    ($AdSyncDDG_Join).ScopeFilter.ScopeConditionList
}

Function Get-ADSyncRuleFreeSlot {
    param ([string]$ruleName)
 
    $ADSyncRulesAndEmptySlots = New-Object Collections.ArrayList
    $allRules = Get-ADSyncRule | Select-Object name, Precedence | Where-Object Precedence -lt 100
    
    $correctSlot = $False

    for ($i = 0; $i -lt 100; $i++) {
        $Entry = New-Object PSObject
        $Entry | Add-Member -NotePropertyName Precedence -NotePropertyValue $i
        $Entry | Add-Member -NotePropertyName Name -NotePropertyValue $null
        if ($i -in $allRules.Precedence) {
            Foreach ($rule in $allRules) {
                If ($rule.Precedence -eq $i) {
                    $Entry.name = $rule.name
                }
            }
        }
        Else {
            $Entry.name = "Free Slot"
        }

        $ADSyncRulesAndEmptySlots.Add($Entry) | out-null
    }
        
    While (!$correctSlot) {
        [array]$slot = $ADSyncRulesAndEmptySlots | Out-GridView  -Title "Please select the precedence for rule `"$ruleName`" by choosing a Free Slot."  -OutputMode Single

        If ($slot.count -eq 1) {
            if ($slot.name -ne "Free Slot") {
                Write-Host "Please select one row out of the Free Slots" -ForegroundColor Red
                #read-Key
            }
            else {
                $correctSlot = $True
            }
        }
        else {
            $correctSlot = $True
        }
    }
    return $slot

}


Write-Host "This dianostic have to be executed on AAD Connect server to have access to AADSync PowerShell Module!" -ForegroundColor Yellow
Read-Key

Clear-Host
$Workloads = "ADSync"
Connect-O365PS $Workloads


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts = get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\UnifiedAudit_$ts"
mkdir $ExportPath -Force | Out-Null

if (!(Get-AADSyncDDGRulesExists )) {
    $CurrentProperty = "Check if no other rules for DDG synchronization already implemented in AAD Connect"
    $CurrentDescription = "Success"
    write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host $CurrentProperty -ForegroundColor Cyan -NoNewline
    Write-Host " - No such rules are created" -ForegroundColor Green
    Start-Sleep -Seconds 5
    new-AADSyncDDGRules
}
else {
    $CurrentProperty = "Check if no other rules for DDG synchronization already implemented in AAD Connect"
    $CurrentDescription = "Already present"
    write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host $CurrentProperty -ForegroundColor Cyan -NoNewline
    Write-Host " - Such rules are already created" -ForegroundColor Red
    Read-Host "Do you want to proceed further?"
    $answer = Get-Choice -OptionsList "Yes", "No"  
    if ($answer -eq 0) {
        new-AADSyncDDGRules
    }
    else {
        Write-Host "You choosed not to go further for implementing the AADConnect rules to sync on-premises DDG to AAD/EXO Contacts" -ForegroundColor Red
        $CurrentProperty = "Some rules to implement AADConnect rules to sync on-premises DDG to AAD/EXO Contacts"
        $CurrentDescription = "Customer choosed to exit"
        write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
        Write-Host "The script will return to main menu"
        read-Key    
        Start-O365TroubleshootersMenu
    }
    
}

#TODO: write all steps in logs
#TODO: check what can be exported in HTML report 
#TODO: track if any error while creating the rules
#TODO: maybe would be better to check if the rules are implemented to show in Report

Read-Key
Start-O365TroubleshootersMenu
