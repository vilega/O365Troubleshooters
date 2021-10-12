Function Get-SpecificRoleMembers {
    param ([string[]]$roles)

    if (!($roles)) {

        $roles = (Get-ManagementRole | Select-Object name | Sort-Object name  |Out-GridView -PassThru -Title "List users that have the role you selected:").Name
    }
    $GetEffectiveUsers = Get-ManagementRoleAssignment -GetEffectiveUsers | Where-Object {(($_.Enabled -eq $True) -and ($roles -match $_.Role))} |`
        Select-Object Role, RoleAssigneeName, RoleAssigneeType, RoleAssignmentDelegationtype, @{ Name = 'Alias(User)'; Expression = {$_.User}}, CustomRecipientWriteScope, CustomConfigWriteScope, RecipientWriteScope, ConfigWriteScope, Identity 
    return $GetEffectiveUsers
}

Function Get-AllUsersWithAllRoles {

    $GetEffectiveUsers = Get-ManagementRoleAssignment -GetEffectiveUsers | Where-Object {($_.Enabled -eq $True)} |`
        Select-Object Role, RoleAssigneeName, RoleAssigneeType, RoleAssignmentDelegationtype, @{ Name = 'Alias(User)'; Expression = {$_.User}}, CustomRecipientWriteScope, CustomConfigWriteScope, RecipientWriteScope, ConfigWriteScope, Identity 
    return $GetEffectiveUsers
}