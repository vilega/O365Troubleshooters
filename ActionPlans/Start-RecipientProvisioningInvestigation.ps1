
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
$EXORecipientbyUPN = Get-Recipient -ResultSize unlimited -IncludeSoftDeletedRecipients | fl userprincipalname, name, WindowsEmailAddress, externaldirectoryObjectid, exchangeguid, guid,  recipienttype, recipienttypedetails, PreviousRecipientTypeDetails

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
	    if(($UPN -match $object.Mail) -and !($object.Mail -eq $null)){
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
	[PSCustomObject]$SearchingInAzureAdUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $MyADUsersCheckObject -TableType "Table"

	$null = $TheObjectToConvertToHTML.Add($SearchingInAzureAdUsers)
	
	
### in AADGroups	

$allAADGroups= Get-AzureADGroup -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$MyADGroupsCheckObject = @()
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD groups"
	foreach($object in $allAADGroups){
	    if(($UPN -match $object.Mail) -and !($object.Mail -eq $null)){
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
	    if(($UPN -match $object.Mail) -and !($object.Mail -eq $null)){
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
	    if($UPN -match $object.SignInName){
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
	    if($UPN -match $object.SignInName){
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
	    if(($UPN -match $object.Mail) -and !($object.Mail -eq $null)){
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
	    if($UPN -match $object.EmailAddress){
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

	### in Recipients	

$allRecipients=Get-Recipient -IncludeInactive -ResultSize unlimited | select DisplayName, Identity, alias, EmailAddresses, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, ExternalDirectoryObjectId, ExchangeGuid, GUID
	
#search on email or alias
$MyRecipientsCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on EXO recipients"
foreach($object in $allRecipients){
	if($UPN -match $object.EmailAddress){
		Write-Host -ForegroundColor Yellow "Found match on property: EmailAddress" 
		Write-Host -ForegroundColor Yellow "`on recipient $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddress"}}
		$MyRecipientsCheckObject = $MyRecipientsCheckObject + $x  
	}      
		
	foreach($proxya in $object.ProxyAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
			Write-Host -ForegroundColor Yellow "`on recipient $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"ProxyAddresses"}}
			$MyRecipientsCheckObject = $MyRecipientsCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on EXO Recipients "
	$MyRecipientsCheckObject = New-Object PSObject
	$MyRecipientsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for EXO Recipients"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInRecipients = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyRecipientsCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInRecipients)


### in MailContacts	

$allMailContacts = Get-MailContact -ResultSize unlimited | select DisplayName, ExternalEmailAddress, EmailAddresses, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, ExternalDirectoryObjectId, GUID

#search on email or externalemail
$MyMailContactsCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on EXO MailContacts"
foreach($object in $allMailContacts){
	if($UPN -match $object.ExternalEmailAddress){
		Write-Host -ForegroundColor Yellow "Found match on property: ExternalEmailAddress" 
		Write-Host -ForegroundColor Yellow "`on MailContact $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"ExternalEmailAddress"}}
		$MyMailContactsCheckObject = $MyMailContactsCheckObject + $x  
	}      
		
	foreach($emailaddress in $object.EmailAddresses){
		if($emailaddress -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on MailContact $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyMailContactsCheckObject = $MyMailContactsCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on EXO Mailcontacts "
	$MyMailContactsCheckObject = New-Object PSObject
	$MyMailContactsCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for EXO MailContacts"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInMailContacts = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyMailContactsCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInMailContacts)


### in Mailusers	

$allMailUsers=Get-Mailuser -ResultSize unlimited | select DisplayName, UserPrincipalName, ExternalEmailAddress, EmailAddresses,  ExternalDirectoryObjectId, ExchangeGuid, GUID
	
#search on email or alias
$MyMailUsersCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on EXO MailUsers"
foreach($object in $allMailUsers){
	if($UPN -match $object.ExternalEmailAddress){
		Write-Host -ForegroundColor Yellow "Found match on property: ExternalEmailAddress" 
		Write-Host -ForegroundColor Yellow "`on MailUser $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"ExternalEmailAddress"}}
		$MyMailUsersCheckObject = $MyMailUsersCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on recipient $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyMailusersCheckObject = $MyMailusersCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on EXO MailUsers "
	$MyMailusersCheckObject = New-Object PSObject
	$MyMailusersCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for EXO Mailusers"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInMailUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyMailUsersCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInMailUsers)


### in sftdeleted Mailusers	

$allSDMailUsers=Get-Mailuser -Softdeleted -ResultSize unlimited | select DisplayName, UserPrincipalName, ExternalEmailAddress, EmailAddresses,  ExternalDirectoryObjectId, ExchangeGuid, GUID
	
#search on email or alias
$MySDMailUsersCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on EXO sfotdeleted MailUsers"
foreach($object in $allSDMailUsers){
	if($UPN -match $object.ExternalEmailAddress){
		Write-Host -ForegroundColor Yellow "Found match on property: ExternalEmailAddress" 
		Write-Host -ForegroundColor Yellow "`on softdeleted MailUser $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"ExternalEmailAddress"}}
		$MySDMailUsersCheckObject = $MySDMailUsersCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on recipient $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MySDMailusersCheckObject = $MySDMailusersCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on EXO softdeleted MailUsers "
	$MySDMailusersCheckObject = New-Object PSObject
	$MySDMailusersCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for EXO softdeleted Mailusers"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInSDMailUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MySDMailUsersCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInSDMailUsers)



### in Mailboxes	

$allMailboxes=Get-MailBox -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on Mailboxes"
foreach($object in $allMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on active mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyMailboxesCheckObject = $MyMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on active mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyMailboxesCheckObject = $MyMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on mailboxes "
	$MyMailboxesCheckObject = New-Object PSObject
	$MyMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInMailboxes)


### in softdeleted Mailboxes	

$allSDMailboxes=Get-MailBox -Softdeleted -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MySDMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on softdeleted Mailboxes"
foreach($object in $allSDMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on softdeleted mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MySDMailboxesCheckObject = $MySDMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on softdeleted mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MySDMailboxesCheckObject = $MySDMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on softdeleted mailboxes "
	$MySDMailboxesCheckObject = New-Object PSObject
	$MySDMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for softdeleted mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInSDMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MySDMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInSDMailboxes)


### in inactive Mailboxes	

$allinMailboxes=Get-MailBox -InactiveMailboxOnly -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyinMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on inactive Mailboxes"
foreach($object in $allinMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on inactive mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyinMailboxesCheckObject = $MyinMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on inactive mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyinMailboxesCheckObject = $MyinMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on inactive mailboxes "
	$MyinMailboxesCheckObject = New-Object PSObject
	$MyinMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for inactive mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingIninMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyinMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingIninMailboxes)


### in Public Folder Mailboxes	

$allPFMailboxes=Get-MailBox -PublicFolder -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyPFMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on Public Folder Mailboxes"
foreach($object in $allMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyPFMailboxesCheckObject = $MyPFMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyPFMailboxesCheckObject = $MyPFMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on PF mailboxes "
	$MyPFMailboxesCheckObject = New-Object PSObject
	$MyPFMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for PF mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInPFMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyPFMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInPFMailboxes)


### in softdeleted Public Folder Mailboxes	

$allSDPFMailboxes=Get-MailBox -PublicFolder -Softdeleted -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MySDPFMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on softdeleted Public Folders Mailboxes"
foreach($object in $allSDPFMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on softdeleted PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MySDPFMailboxesCheckObject = $MySDPFMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on softdeleted PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MySDPFMailboxesCheckObject = $MySDPFMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on softdeleted PF mailboxes "
	$MySDPFMailboxesCheckObject = New-Object PSObject
	$MySDPFMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for softdeleted PF mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInSDPFMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MySDPFMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInSDPFMailboxes)


### in inactive Public Folder Mailboxes	

$allinPFMailboxes=Get-MailBox -PublicFolder -InactiveMailboxOnly -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyinPFMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on inactive PF Mailboxes"
foreach($object in $allinPFMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on inactive PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyinPFMailboxesCheckObject = $MyinPFMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on inactive PF mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyinPFMailboxesCheckObject = $MyinPFMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on inactive PF mailboxes "
	$MyinPFMailboxesCheckObject = New-Object PSObject
	$MyinPFMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for inactive PF mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingIninPFMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyinPFMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingIninPFMailboxes)



### in Group Mailboxes	

$allGRMailboxes=Get-MailBox -GroupMailbox -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyGRMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on Group Mailboxes"
foreach($object in $allMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on active group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyGRMailboxesCheckObject = $MyGRMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on active group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyGRMailboxesCheckObject = $MyGRMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on group mailboxes "
	$MyGRMailboxesCheckObject = New-Object PSObject
	$MyGRMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for group mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInGRMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyGRMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInGRMailboxes)


### in softdeleted Group Mailboxes	

$allSDGRMailboxes=Get-MailBox -GroupMailbox -Softdeleted -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MySDGRMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on softdeleted Group Mailboxes"
foreach($object in $allSDGRMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on softdeleted group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MySDGRMailboxesCheckObject = $MySDGRMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on softdeleted group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MySDGRMailboxesCheckObject = $MySDGRMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on softdeleted group mailboxes "
	$MySDGRMailboxesCheckObject = New-Object PSObject
	$MySDGRMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for softdeleted group mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInSDGRMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MySDGRMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInSDGRMailboxes)


### in inactive Group Mailboxes	

$allinGRMailboxes=Get-MailBox -GroupMailbox -InactiveMailboxOnly -ResultSize unlimited | select DisplayName, UserPrincipalName, EmailAddresses, ExternalDirectoryObjectId, ExchangeGuid, GUID

#search on email or upn
$MyinGRMailboxesCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Searching on inactive Group Mailboxes"
foreach($object in $allinGRMailboxes){
	if($UPN -match $object.UserPrincipalName){
		Write-Host -ForegroundColor Yellow "Found match on property: UserPrincipalName" 
		Write-Host -ForegroundColor Yellow "`on inactive group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"UserPrincipalName"}}
		$MyinGRMailboxesCheckObject = $MyinGRMailboxesCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on inactive group mailbox $($object.DisplayName) having Guid $($object.Guid) and ExternalDirectoryObjectId $($object.ExternalDirectoryObjectId)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyinGRMailboxesCheckObject = $MyinGRMailboxesCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on inactive group mailboxes "
	$MyinGRMailboxesCheckObject = New-Object PSObject
	$MyinGRMailboxesCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for inactive group mailboxes"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingIninGRMailboxes = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyinGRMailboxesCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingIninGRMailboxes)


### in Mail-enabled Public Folders	

$allMailEnabledPF=Get-MailPublicFolder -ResultSize unlimited | select DisplayName, WindowsEmailAddress, EmailAddresses, GUID

	
#search on email or WindowsEmailAddress
$MyMailEnabledPFCheckObject = @()
$FoundExistence=$false
Write-Host -ForegroundColor Magenta "Mail-enabled Public Folders"
foreach($object in $allMailEnabledPF){
	if($UPN -match $object.WindowsEmailAddress){
		Write-Host -ForegroundColor Yellow "Found match on property: WindowsEmailAddress" 
		Write-Host -ForegroundColor Yellow "`on mail-enabled PF $($object.DisplayName) having Guid $($object.Guid)"
		$FoundExistence=$true 
		$x = $object | Select-Object DisplayName, Guid, @{name="Attribute"; expression={"WindowsEmailAddress"}}
		$MyMailEnabledPFCheckObject = $MyMailEnabledPFCheckObject + $x  
	}      
		
	foreach($proxya in $object.EmailAddresses){
		if($proxya -match $UPN){
			Write-Host -ForegroundColor Yellow "Found match on property: EmailAddresses" 
			Write-Host -ForegroundColor Yellow "`on mail-enabled PF $($object.DisplayName) having Guid $($object.Guid)" 
			$FoundExistence=$true
			$x = $object | Select-Object DisplayName, Guid, ExternalDirectoryObjectId, @{name="Attribute"; expression={"EmailAddresses"}}
			$MyMailEnabledPFCheckObject = $MyMailEnabledPFCheckObject + $x  
		}      
	}
}
if(!$FoundExistence){
	Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on mail-enabled PF"
	$MyMailEnabledPFCheckObject = New-Object PSObject
	$MyMailEnabledPFCheckObject | Add-Member -NotePropertyName Objects -NotePropertyValue "Found no objects matching value"
}

[string]$SectionTitle = "Searching for mail-enabled public folders"
[string]$Description = "Check for multiple conflicting objects"

[PSCustomObject]$SearchingInMailEnabledPF = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
-DataType "ArrayList" -EffectiveDataArrayList $MyMailEnabledPFCheckObject -TableType "Table"

$null = $TheObjectToConvertToHTML.Add($SearchingInMailEnabledPF)





	### Reporting out
	[string]$FilePath = $ExportPath + "\RecipientProvisioning_Report.html"
    Export-ReportToHTML -FilePath $FilePath -PageTitle "Recipient Provisioning Issues Report" -ReportTitle "Recipient Provisioning Issues Report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
