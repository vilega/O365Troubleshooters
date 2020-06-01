Function Get-SpecificRoleMembers {
    param ([string[]]$roles)

    if (!($roles)) {

        $roles = (Get-ManagementRole | Select-Object name | Sort-Object name  |Out-GridView -PassThru -Title "List users that have the role you selected:").Name
    }
    $GetEffectiveUsers = Get-ManagementRoleAssignment -GetEffectiveUsers | Where-Object {(($_.Enabled -eq $True) -and ($roles -match $_.Role))} |`
        Select-Object Role, RoleAssigneeName, RoleAssigneeType, RoleAssignmentDelegationtype, User, CustomRecipientWriteScope, CustomConfigWriteScope, RecipientWriteScope, ConfigWriteScope, Identity |`
        export-csv "$ExportPath\RoleMembers_$ts.csv" -NoTypeInformation 
    Write-Host "The list of user who have selected roles assigned was exported to $global:ExportPath\RoleMembers_$ts.csv"
    return $GetEffectiveUsers
}

Function Get-AllUsersWithAllRoles {

    Get-ManagementRoleAssignment -GetEffectiveUsers | Where-Object {($_.Enabled -eq $True)} |`
         Select-Object Role, RoleAssigneeName, RoleAssigneeType, RoleAssignmentDelegationtype, User, CustomRecipientWriteScope, CustomConfigWriteScope, RecipientWriteScope, ConfigWriteScope, Identity |`
         export-csv "$global:ExportPath\ManagementRoleAssignmentUsers_$ts.csv" -NoTypeInformation
    Write-Host "Export all users with all the roles assigned to the file: $global:ExportPath\ManagementRoleAssignmentUsers_$ts.csv"
}