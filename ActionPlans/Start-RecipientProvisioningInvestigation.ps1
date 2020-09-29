
Function Search-RecipientObject
{
    param( 
        [int][Parameter(Mandatory=$true)] $DaysToSearch,
        [string[]][Parameter(Mandatory=$false)] $OperationsToSearch,
        [string][Parameter(Mandatory=$false)] $userIds)
      
    $DaysToSearch=10
    if (!([string]::IsNullOrEmpty($userIds)))
    {
        $UnifiedAuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch - -UserIds $userIds -SessionCommand ReturnLargeSet 
    }
    else
    {
        $UnifiedAuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).addDays(-$DaysToSearch) -EndDate (Get-Date) -Operations $OperationsToSearch  -SessionCommand ReturnLargeSet 
    }
  
    return $UnifiedAuditLogs

}


Clear-Host
#import-module C:\Work\Projects\PS\GitHubStuff\O365Troubleshooters\O365Troubleshooters.psm1 -Force
#Set-GlobalVariables
#Start-O365TroubleshootersMenu

$Workloads = "exo", "msol", "azuread"
Connect-O365PS $Workloads
$CurrentProperty = "Connecting to: $Workloads"
$CurrentDescription = "Success"
write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription
$ts= get-date -Format yyyyMMdd_HHmmss
$ExportPath = "$global:WSPath\RecipientProvisioningInvestigation_$ts"
mkdir $ExportPath -Force |out-null

# Obtain UPN of the affected user
$upn = Get-ValidEmailAddress("UserPrincipalName")


# Lookup MSOl object by UPN
$MSOLUserbyUPN = Get-MsolUser -all | ?{$_.UserPrincipalName -match $upn}
$MSOLContactbyUPN = Get-MsolContact -all | ?{$_.UserPrincipalName -match $upn}
$MSOLGroupbyUPN = Get-MsolGroup -all | ?{$_.UserPrincipalName -match $upn}

#Lookup AzureAd object by UPN
$AADUserbyUPN = Get-AzureADUser -all:$true | ?{$_.UserPrincipalName -match $upn}
$AADContactbyUPN = Get-AzureADContact -all:$true | ?{$_.UserPrincipalName -match $upn}
$AADGroupbyUPN = Get-AzureADGroup -all:$true | ?{$_.UserPrincipalName -match $upn}

# Search EXO AD object by UPN
$EXOADUserbyUPN = Get-User -ResultSize unlimited | ?{$_.UserPrincipalName -match $upn}
$EXOADObjectbyUPNFound = Get-Recipient -ResultSize unlimited -IncludeSoftDeletedRecipients | fl userprincipalname, name, WindowsEmailAddress, externaldirectoryObjectid, exchangeguid, guid,  recipienttype, recipienttypedetails, PreviousRecipientTypeDetails

# Export by UPN found info
$MSOLUserbyUPN | export-CliXml -Depth 3 $path\"MsolUser.xml"
$MSOLContactbyUPN | export-CliXml -Depth 3 $path\"MsolContact.xml"
$MSOLGroupbyUPN | export-CliXml -Depth 3 $path\"MsolGroup.xml"
$AADUserbyUPN | export-CliXml -Depth 3 $path\"AADUser.xml"
$AADContactbyUPN | export-CliXml -Depth 3 $path\"AADContact.xml"
$AADGroupbyUPN | export-CliXml -Depth 3 $path\"AADGroup.xml"
$EXOADUserbyUPN | export-CliXml -Depth 3 $path\"EXOUser.xml"
Get-Recipient $EXOADUserbyUPN.ExternalDirectoryObjectId | export-CliXml -Depth 3 $path\"EXORecipient.xml"

# Output objects definition

[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()


# Prepare-ObjectForHTMLReport

# Check if UPN value is already in use on different object as other property

## Check in AzureAD
### in AADUsers	

$allAADUsers= Get-AzureADUser -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyADUsersCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD users"
	foreach($object in $allAADUsers){
	    if($UPN -eq $object.Mail){
	        Write-Host -ForegroundColor Yellow "Found match on property: Mail" 
			Write-Host -ForegroundColor Yellow "`on AzureAD user $($object.DisplayName) having ObjectId $($object.ObjectId)"
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"Mail"}}
			$MyADUsersCheckObject = $MyADUsersCheckObject + $x
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyADUsersCheckObject = $MyADUsersCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD users "
			
		$MyADGroupsCheckObject = New-Object PSObject
		$MyADGroupsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}
	
	[string]$SectionTitle = "Searching for AzureADUsers"
	[string]$Description = "Check for multiple conflicting objects"
	#select object id unique, count
	[PSCustomObject]$SearchingInAzureAdUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyADUsersCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInAzureAdUsers)
	
	
### in AADGroups	

$allAADGroups= Get-AzureADGroup -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyADGroupsCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD groups"
	foreach($object in $allAADGroups){
	    if($UPN -eq $object.Mail){
	        Write-Host -ForegroundColor Yellow "Found match on property: Mail" 
			Write-Host -ForegroundColor Yellow "`on AzureAD group $($object.DisplayName) having ObjectId $($object.ObjectId)"
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"Mail"}}
			$MyADGroupsCheckObject = $MyADGroupsCheckObject + $x 
		}  
		    
	        foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyADGroupsCheckObject = $MyADGroupsCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD groups "
			
		$MyADGroupsCheckObject = New-Object PSObject
		$MyADGroupsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
		#$MyADGroupsCheckObject.GetType()
		#$MyADGroupsCheckObject | Get-Member
	}

	[string]$SectionTitle = "Searching for AzureADGroups"
	[string]$Description = "Check for multiple conflicting objects"
	#select object id unique, count
	[PSCustomObject]$SearchingInAzureADGroups = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyADGroupsCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInAzureADGroups)

### in AADContacts	

$allAADContacts= Get-AzureADContact -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyADContactsCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD contacts"
	foreach($object in $allAADContacts){
	    if($UPN -eq $object.Mail){
	        Write-Host -ForegroundColor Yellow "Found match on property: Mail" 
			Write-Host -ForegroundColor Yellow "`on AzureAD contact $($object.DisplayName) having ObjectId $($object.ObjectId)"
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"Mail"}}
			$MyADContactsCheckObject = $MyADContactsCheckObject + $x 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyADContactsCheckObject = $MyADContactsCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD contacts "
		$MyADContactsCheckObject = New-Object PSObject
		$MyADContactsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}

	[string]$SectionTitle = "Searching for AzureADContacts"
	[string]$Description = "Check for multiple conflicting objects"
	
	[PSCustomObject]$SearchingInAzureADContacts = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyADContactsCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInAzureADContacts)

## Check in MSOL
### in MSOLUsers	

$allMSOLUsers= Get-MsolUser -All | select DisplayName,SignInName,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyMSOLUsersCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL users"
	foreach($object in $allMSOLUsers){
	    if($UPN -eq $object.SignInName){
	        Write-Host -ForegroundColor Yellow "Found match on property: SignInName" 
			Write-Host -ForegroundColor Yellow "`on MSOL user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"SignInName"}}
			$MyMSOLUsersCheckObject = $MyMSOLUsersCheckObject + $x 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyMSOLUsersCheckObject = $MyMSOLUsersCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL users "
		$MyMSOLUsersCheckObject = New-Object PSObject
		$MyMSOLUsersCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}

	[string]$SectionTitle = "Searching for MSOL users"
	[string]$Description = "Check for multiple conflicting objects"
	
	[PSCustomObject]$SearchingInMSOLUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyMSOLUsersCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInMSOLUsers)

### in deleted MSOLUsers	

$allMSOLDeletedUsers= Get-MsolUser -All -ReturnDeletedUsers | select DisplayName,SignInName,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyMSOLDeletedUsersCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on deleted MSOL users"
	foreach($object in $allMSOLDeletedUsers){
	    if($UPN -eq $object.SignInName){
	        Write-Host -ForegroundColor Yellow "Found match on property: SignInName" 
			Write-Host -ForegroundColor Yellow "`on MSOL deleted user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"SignInName"}}
			$MyMSOLDeletedUsersCheckObject = $MyMSOLDeletedUsersCheckObject + $x 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL deleted user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyMSOLDeletedUsersCheckObject = $MyMSOLDeletedUsersCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL deleted users "
		$MyMSOLDeletedUsersCheckObject = New-Object PSObject
		$MyMSOLDeletedUsersCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}

	[string]$SectionTitle = "Searching for MSOL deleted users"
	[string]$Description = "Check for multiple conflicting objects"
	
	[PSCustomObject]$SearchingInMSOLDeletedUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyMSOLDeletedUsersCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInMSOLDeletedUsers)


### in MSOLGroups	

$allMSOLGroups= Get-MSOLGroup -All:$true | select DisplayName,EmailAddress,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyMSOLGroupsCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL groups"
	foreach($object in $allMSOLGroups){
	    if($UPN -eq $object.EmailAddress){
	        Write-Host -ForegroundColor Yellow "Found match on property: EmailAddress" 
			Write-Host -ForegroundColor Yellow "`on MSOL group $($object.DisplayName) having ObjectId $($object.ObjectId)"
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"EmailAddress"}}
			$MyMSOLGroupsCheckObject = $MyMSOLGroupsCheckObject + $x  
		}  
		    
	        foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyMSOLGroupsCheckObject = $MyMSOLGroupsCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL groups "
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL users "
		$MyMSOLGroupsCheckObject = New-Object PSObject
		$MyMSOLGroupsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}

	[string]$SectionTitle = "Searching for MSOL users"
	[string]$Description = "Check for multiple conflicting objects"
	
	[PSCustomObject]$SearchingInMSOLGroups = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyMSOLGroupsCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInMSOLGroups)

### in MSOLContacts	

$allMSOLContacts= Get-MSOLContact -All:$true | select DisplayName,EmailAddress,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyMSOLContactsCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL contacts"
	foreach($object in $allMSOLContacts){
	    if($UPN -eq $object.EmailAddress){
	        Write-Host -ForegroundColor Yellow "Found match on property: EmailAddress" 
			Write-Host -ForegroundColor Yellow "`on MSOL contact $($object.DisplayName) having ObjectId $($object.ObjectId)"
			$FoundExistence=$true 
			$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"EmailAddress"}}
			$MyMSOLContactsCheckObject = $MyMSOLContactsCheckObject + $x  
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
				$FoundExistence=$true
				$x = $object | Select-Object DisplayName, ObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
				$MyMSOLContactsCheckObject = $MyMSOLContactsCheckObject + $x
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL contacts "
		$MyMSOLContactsCheckObject = New-Object PSObject
		$MyMSOLContactsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
	}

	[string]$SectionTitle = "Searching for MSOL users"
	[string]$Description = "Check for multiple conflicting objects"
	
	[PSCustomObject]$SearchingInMSOLContacts = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyMSOLContactsCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInMSOLContacts)

	
	## Check in EXO

	


	### Reporting out
	[string]$FilePath = $ExportPath + "\RecipientProvisioning_Report.html"
    Export-ReportToHTML -FilePath $FilePath -PageTitle "Recipient Provisioning Issues Report" -ReportTitle "Recipient Provisioning Issues Report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
