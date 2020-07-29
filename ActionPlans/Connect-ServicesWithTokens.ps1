function Get-Token {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("EXO","AzureGraph","AIPService")]
        [string]
        $Service
    )

    switch ($Service) {
        exo {
            # EXO Powershell Client ID
            $clientId = "a0c73c16-a7e3-4564-9a95-2bdf47383716" 
            # Set redirect URI for PowerShell
            $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
            # Set Resource URI to EXO endpoint
            $resourceId = "https://outlook.office365.com"
            # Set Authority to Azure AD Tenant
            $authority = "https://login.microsoftonline.com/common"

        }
        AzureGraph {
            # Azure PowerShell Client ID
            $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
            # Redirect URI for PowerShell
            $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
            # Resource URI to Graph endpoint
            $resourceId = "https://graph.windows.net"
            # Authority to Azure AD Tenant
            $authority = "https://login.microsoftonline.com/common"           
        }

        AIPService {
            # AIP PowerShell Client ID
            $clientId='90f610bf-206d-4950-b61d-37fa6fd1b224';
            # Resource URI to AADRM endpoint
            $resourceId = 'https://api.aadrm.com/';
             # Authority to Azure AD Tenant
            $authority = "https://login.microsoftonline.com/common";
        }
    }
    
    # AuthenticationContext points to Azure AD
    $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($authority);

    # Get token silently
    $UserIdentifier = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier($global:credentials.userName, "OptionalDisplayableId")
    $authResult = $authContext.AcquireTokenSilentAsync($resourceId, $clientId,$UserIdentifier)
    while ($authResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500}
    if (!($authResult.IsFaulted -eq $false)) 
    {
        switch ($Result.Exception.InnerException.ErrorCode) 
        {
            failed_to_acquire_token_silently 
            {
                $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($authority);
                $userName = $global:credentials.userName
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:credentials.password))
                $userCreds = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential($userName, $password)
                $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceId, $clientId, $userCreds)
                while ($authResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500}
                $Result = $authResult.Result
            }
            multiple_matching_tokens_detected 
            {
                # It shouldn't reach this point as we are requeting the token for a spefic service & UPN
                $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($authority);
                $userName = $global:credentials.userName
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:credentials.password))
                $userCreds = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential($userName, $password)
                $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceId, $clientId, $userCreds)
                while ($authResult.IsCompleted -ne $true) { Start-Sleep -Milliseconds 500}
            }
            default 
            {
                $CurrentProperty = "Get token silently"
                $CurrentDescription = "Unknown error $($Result.Exception.InnerException.ErrorCode)"
                Write-Host "`n$CurrentDescription" -ForegroundColor Red
                write-log -Function "Get-Token" -Step $CurrentProperty -Description $CurrentDescription

            }
        }
    }
    return $authResult.Result
}

function Get-TokenFromCache
{
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("EXO","AzureGraph","AIPService")]
        [string]
        $Service
    )
      
    switch ($Service) 
    {
        exo 
        {
            $resourceId = "https://outlook.office365.com"
        }
        AzureGraph 
        {
            # Resource URI to Graph endpoint
            $resourceId = "https://graph.windows.net"
        }

        AIPService 
        {
            # Resource URI to AADRM endpoint
            $resourceId = 'https://api.aadrm.com/';
        }
    }

    $cache = [Microsoft.IdentityModel.Clients.ActiveDirectory.TokenCache]::DefaultShared
    return $Cache.ReadItems() | Where-Object {($_.DisplayableId -eq $global:credentials.userName) -and ($_.Resource -eq $resourceID)}
   
}