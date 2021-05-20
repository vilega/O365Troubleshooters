# Function to validate if required custom rules exists and to create them 
Function new-AADSyncDDGRules {

    # select local AD connector
    $selectConnector = Get-ADSyncConnector | Where-Object { $_.Name -notlike "*onmicrosoft*" } | Select-Object name, identifier | Out-GridView -Title "Please select the connector to your local Active Directory!" -OutputMode Single
    if ($selectConnector)
    {
        [string]$SectionTitle = "On-premises AD Connector used by the created rules"
        [string]$Description = "The rules will be created using the selected connector: $($selectConnector.name)"
        [PSCustomObject]$OnPremisesADConnectorSelectedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring "See next sections for rule creations!"
        $null = $TheCollectionToConvertToHTML.Add($OnPremisesADConnectorSelectedHTML)

    #region Create ADSync PROVISION rule
    $ruleName = 'Custom In from AD - Dynamic Distribtution Group - Provision'
    if (!(Get-ADSyncRule | Where-Object { ($_.Name -eq $ruleName) -and ($_.Connector -eq $selectConnector.identifier.guid) })) {

        $slot = Get-ADSyncRuleFreeSlot -ruleName $ruleName 
        if ($slot.count -eq 0) {
            Write-Host "You select CANCEL so the rule `"$ruleName`" won't be created"
            Read-Key
            #TODO: write this on the report

            [string]$SectionTitle = "`"$ruleName`" - creation"
            [string]$Description = "The rule `"$ruleName`" has not been created using selected connector: $($selectConnector.name) because no slot was selected"
            [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring ""
            $null = $TheCollectionToConvertToHTML.Add($RuleHTML)
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
                -Connector $selectConnector.identifier.guid `
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

            $CurrentProperty = "New-AADSyncDDGRules"
            $CurrentDescription = "Rule `"$ruleName`" has been created"
            Write-Host $CurrentDescription -ForegroundColor Green
            write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
            
            [string]$SectionTitle = "`"$ruleName`" - creation"
            [string]$Description = "The rule `"$ruleName`" has been succesfully created using selected connector: $($selectConnector.name)"
            [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring ""
            $null = $TheCollectionToConvertToHTML.Add($RuleHTML)
            
            Read-Key    


        }
    }
    else {
        Write-Host "Rule `"$ruleName`" already exist. If you want to re-create it, you can delete it from AAD Connect and re-run the script to re-created with the latest version"
        $CurrentProperty = "New-AADSyncDDGRules"
        $CurrentDescription = "rule `"$ruleName`" was already created"
        write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
        
        [string]$SectionTitle = "`"$ruleName`" - creation"
        [string]$Description = "The rule `"$ruleName`" has not been succesfully created using selected connector: $($selectConnector.name) because already exists."
        [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring ""
        $null = $TheCollectionToConvertToHTML.Add($RuleHTML)

        Read-Key    
    }
    #endregion Create the PROVISION rule


    #region Create ADSync JOIN rule (Has All Attribute Transformations)
    $ruleName = 'Custom In from AD - Dynamic Distribution Group - Join'
    if (!(Get-ADSyncRule | Where-Object { ($_.Name -eq $ruleName) -and ($_.Connector -eq $selectConnector.identifier.guid) })) {

        $slot = Get-ADSyncRuleFreeSlot -ruleName $ruleName 
        if ($slot.count -eq 0) {
            Write-Host "You select CANCEL so the rule `"$ruleName`" won't be created"
            #TODO: write this on the report
            #TODO: write to the log file
            [string]$SectionTitle = "`"$ruleName`" - creation"
            [string]$Description = "The rule `"$ruleName`" has not been created using selected connector: $($selectConnector.name) because no slot was selected"
            [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring ""
            $null = $TheCollectionToConvertToHTML.Add($RuleHTML)

            Read-Key

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
                -Connector $selectConnector.identifier.guid  `
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
                -OutVariable syncRule  | Out-Null


            Add-ADSyncAttributeFlowMapping  `
                -SynchronizationRule $syncRule[0] `
                -Source @('unauthOrig') `
                -Destination 'unauthOrig' `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -OutVariable syncRule  | Out-Null


            #endregion Create Attribute Flow Mappings

            #region Create Scope Condition to match all Dynamic Distribtuion object type
            New-Object  `
                -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
                -ArgumentList 'msExchRecipientDisplayType', '3', 'EQUAL' `
                -OutVariable condition0  | Out-Null

            Add-ADSyncScopeConditionGroup  `
                -SynchronizationRule $syncRule[0] `
                -ScopeConditions @($condition0[0]) `
                -OutVariable syncRule  | Out-Null
            #endregion Create Scope Condition to match all Dynamic Distribtuion object type


            # Add the rule in the AAD Connect engine
            Add-ADSyncRule -SynchronizationRule $syncRule[0] | Out-Null
            Remove-Variable syncRule -Force
            $CurrentProperty = "New-AADSyncDDGRules"
            $CurrentDescription = "Rule `"$ruleName`" has been created"
            Write-Host $CurrentDescription -ForegroundColor Green
            write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
            
            [string]$SectionTitle = "`"$ruleName`" - creation"
            [string]$Description = "The rule `"$ruleName`" has been succesfully created using selected connector: $($selectConnector.name)"
            [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "String" -EffectiveDatastring ""
            $null = $TheCollectionToConvertToHTML.Add($RuleHTML)
            
            

            Read-Key    
        }
    }
    else {
        Write-Host "Rule `"$ruleName`" already exist. If you want to re-create it, you can delete it from AAD Connect and re-run the script to re-created with the latest version"
        $CurrentProperty = "New-AADSyncDDGRules"
        $CurrentDescription = "rule `"$ruleName`" was already created"
        write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 

        [string]$SectionTitle = "`"$ruleName`" - creation"
        [string]$Description = "The rule `"$ruleName`" has not been succesfully created using selected connector: $($selectConnector.name) because already exists."
        [PSCustomObject]$RuleHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring ""
        $null = $TheCollectionToConvertToHTML.Add($RuleHTML)

        Read-Key    
    }
}
else {
    #TODO: write in report
    $CurrentProperty = "selectConnector"
    $CurrentDescription = "No connector was selected!"
    write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 
    Write-Host $CurrentDescription

    [string]$SectionTitle = "On-premises AD Connector used by the created rules"
    [string]$Description = "No On-premises AD Connector was selected."
    [PSCustomObject]$OnPremisesADConnectorSelectedHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Red" -Description $Description -DataType "String" -EffectiveDatastring "No rules will be created!"
    $null = $TheCollectionToConvertToHTML.Add($OnPremisesADConnectorSelectedHTML)

    read-Key    
}
}

# Check if Admin previously created similar rules (not used now as a different check was implemented)
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

# Show free slots for rules to administrator and allow him to chose the position where the rules will be created
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

# Inform about the prerequisites 
Write-Host "This dianostic have to be executed on AAD Connect server to have access to AADSync PowerShell Module!" -ForegroundColor Yellow
Read-Key

Clear-Host

# Import Module ADSync 
$Workloads = "ADSync"
Connect-O365PS -O365Service $Workloads -requireCredentials $false


$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to workloads" -Step $CurrentProperty -Description $CurrentDescription 
    
$ts = get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\SyncDDGasContactwithAADConnect_$ts"
mkdir $ExportPath -Force | Out-Null


##Initialize HTML Object
[System.Collections.ArrayList]$TheCollectionToConvertToHTML = @()

[string]$SectionTitle = "Information"
[string]$article1='<a href="https://docs.microsoft.com/previous-versions/office/exchange-server-2010/jj150422(v=exchg.141)?redirectedfrom=MSDN" target="_blank">Configure Dynamic Distribution Groups in a Hybrid Deployment</a>'
[string]$article2='<a href="https://answers.microsoft.com" target="_blank">Creating AAD Connect rules to synchronize on-premises Dynamic Distribution Groups as Exchange Online Contacts</a>'
[string]$Description = "In a Hybrid Exchange environment, the on-premises Dynamic Distribution Groups (DDGs) are not synced in Azure AD / Exchange Online and this is by design. To workaround this limitation Microsoft recommends to create Exchange Online contacts for every on-premises dynamic distribution group. See $article1. To automatically maintain this Exchange Online contacts we can implement some custom AAD Connect rules without touching any default ones. For more information see article $article2."
[PSCustomObject]$InformatioHTML = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "String" -EffectiveDatastring "Please check bellow for what was implemented!"
$null = $TheCollectionToConvertToHTML.Add($InformatioHTML)

# Write current step in the log
$CurrentProperty = "new-AADSyncDDGRules"
$CurrentDescription = "Start"
write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 

# Call the function to create the AAD Connect rules
new-AADSyncDDGRules

# Write current step in the log
$CurrentProperty = "new-AADSyncDDGRules"
$CurrentDescription = "Finished"
write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 

# Creating the Final report with all sections previously created
[string]$FilePath = $ExportPath + "\AADConnectSyncDDGasContacts.html"
Export-ReportToHTML -FilePath $FilePath -PageTitle "Creating AAD Connect rules to synchronize on-premises Dynamic Distribution Groups as Exchange Online Contacts" -ReportTitle "Implementing AAD Connect rules to synchronize on-premises Dynamic Distribution Groups as Exchange Online Contacts" -TheObjectToConvertToHTML $TheCollectionToConvertToHTML

# Write current step in the log
$CurrentProperty = "HTML report which contains information about creation of the AAD Connect rules can be found here: $ExportPath "
$CurrentDescription = "The script will return to main menu!"
write-log -Function "Start-SyncDDGasContactwithAADConnect" -Step $CurrentProperty -Description $CurrentDescription 

# Ask enduser for opening the HTMl report
Write-Host "`nReport was exported in the following location: $ExportPath" -ForegroundColor Cyan 
$OpenHTMLfile=Read-Host "Do you wish to open HTML report file now?`nType Y(Yes) to open or N(No) to exit!"
if ($OpenHTMLfile.ToLower() -eq "y")
{
    # Open report with default browswer
    Write-Host "Opening report with default browswer." -ForegroundColor Cyan
    Start-Process $FilePath
    
    Write-Host "The script will return to main menu!" -ForegroundColor Cyan
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu

}
elseif($OpenHTMLfile.ToLower() -eq "n")
{
    Write-Host "The script will return to main menu!" -ForegroundColor Cyan
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
}
else {
    Write-Host "You didn't provide an expected input!" -ForegroundColor Yellow
    Write-Host "The script will return to main menu!" -ForegroundColor Cyan
    Read-Key
    # Go back to the main menu
    Start-O365TroubleshootersMenu
}
