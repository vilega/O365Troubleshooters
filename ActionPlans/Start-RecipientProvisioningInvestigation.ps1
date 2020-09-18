
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
import-module C:\Work\Projects\PS\GitHubStuff\O365Troubleshooters\O365Troubleshooters.psm1
Set-GlobalVariables
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

$null = $TheObjectToConvertToHTML.Add($TheCommand)



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
			$x = $object | Select-Object DisplayName, ObjectId
			$MyADUsersCheckObject = $MyADUsersCheckObject + $x
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
		Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD users "
	}
	
	[string]$SectionTitle = "Searching for AzureADUsers"

    [string]$Description = "Check for multiple conflicting objects"

	[PSCustomObject]$SearchingInAzureAdUsers = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Green" -Description $Description `
	-DataType "ArrayList" -EffectiveDataArrayList $MyADUsersCheckObject -TableType "List"

	$null = $TheObjectToConvertToHTML.Add($SearchingInAzureAdUsers)
	

### in AADGroups	

$allAADGroups= Get-AzureADGroup -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD groups"
	foreach($object in $allAADGroups){
	    if($UPN -eq $object.Mail){
	        Write-Host -ForegroundColor Yellow "Found match on property: Mail" 
	        Write-Host -ForegroundColor Yellow "`on AzureAD group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
		}  
		    
	        foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD groups "
	}

### in AADContacts	

$allAADContacts= Get-AzureADContact -All:$true | select DisplayName,mail,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on AzureAD contacts"
	foreach($object in $allAADContacts){
	    if($UPN -eq $object.Mail){
	        Write-Host -ForegroundColor Yellow "Found match on property: Mail" 
	        Write-Host -ForegroundColor Yellow "`on AzureAD contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on AzureAD contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on AzureAD contacts "
	}

## Check in MSOL
### in MSOLUsers	

$allMSOLUsers= Get-MsolUser -All | select DisplayName,SignInName,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL users"
	foreach($object in $allMSOLUsers){
	    if($UPN -eq $object.SignInName){
	        Write-Host -ForegroundColor Yellow "Found match on property: SignInName" 
	        Write-Host -ForegroundColor Yellow "`on MSOL user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL users "
	}

### in deleted MSOLUsers	

$allMSOLDeletedUsers= Get-MsolUser -All -ReturnDeletedUsers | select DisplayName,SignInName,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on deleted MSOL users"
	foreach($object in $allMSOLDeletedUsers){
	    if($UPN -eq $object.SignInName){
	        Write-Host -ForegroundColor Yellow "Found match on property: SignInName" 
	        Write-Host -ForegroundColor Yellow "`on MSOL deleted user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL deleted user $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL deleted users "
	}


### in MSOLGroups	

$allMSOLGroups= Get-MSOLGroup -All:$true | select DisplayName,EmailAddress,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL groups"
	foreach($object in $allMSOLGroups){
	    if($UPN -eq $object.EmailAddress){
	        Write-Host -ForegroundColor Yellow "Found match on property: EmailAddress" 
	        Write-Host -ForegroundColor Yellow "`on MSOL group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
		}  
		    
	        foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL group $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL groups "
	}

### in MSOLContacts	

$allMSOLContacts= Get-MSOLContact -All:$true | select DisplayName,EmailAddress,ProxyAddresses,ObjectId
	
	#search on email or alias
	$FoundExistence=$false
    Write-Host -ForegroundColor Magenta "Searching on MSOL contacts"
	foreach($object in $allMSOLContacts){
	    if($UPN -eq $object.EmailAddress){
	        Write-Host -ForegroundColor Yellow "Found match on property: EmailAddress" 
	        Write-Host -ForegroundColor Yellow "`on MSOL contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	    }      
	    	
	    foreach($proxya in $object.ProxyAddresses){
	        if($proxya -match $UPN){
	            Write-Host -ForegroundColor Yellow "Found match on property: ProxyAddress (Alias)" 
	            Write-Host -ForegroundColor Yellow "`on MSOL contact $($object.DisplayName) having ObjectId $($object.ObjectId)" 
	            $FoundExistence=$true
	        }      
	    }
	}
	if(!$FoundExistence){
	    Write-host -ForegroundColor Red """$($UPN)"" Email Address not found on MSOL contacts "
	}

	## Check in EXO

	