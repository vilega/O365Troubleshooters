

#region Common Script Blocks
    # Getting Credentials script block 
    $Global:UserCredential = {
            Write-Host "`nPlease enter Office 365 Global Admin credentials:" -ForegroundColor Cyan
            $Global:O365Cred = Get-Credential
    }

    # Credential Validation block
    $Global:CredentialValidation = { 
            If (!([string]::IsNullOrEmpty($errordescr))-and !([string]::IsNullOrEmpty($global:error[0]))) {
                Write-Host "`nYou are NOT connected succesfully to $Global:banner. Please verify your credentials." -ForegroundColor Yellow
                $CurrentDescription = "`""+$CurrentError+"`""
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                #&$Global:UserCredential
            }
    }

    # Displaying connection status
    $Global:DisplayConnect = {
            If ($errordescr -ne $null) {
                Write-Host "`nYou are NOT connected succesfully to $Global:banner" -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description "You are NOT connected succesfully to $Global:banner"
                Write-Host "`nThe script will now exit." -ForegroundColor Red
                Read-Host
                exit
            }
            else {
                Write-Host "`nYou are connected succesfully to $Global:banner" -ForegroundColor Green
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description "You are connected succesfully to $Global:banner"
            }
    }

     # SPO connection script block:
     $Global:SPOConnectBlock = {
                        
                        $Global:Error.Clear();
                        $Global:banner = "SharePoint Online PowerShell"
                        $try++
                        # Import SPS Online PS module
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
                        # Creating a new SPS Online PS Session
                        $Global:UrlSharepoint = "https://$DomainHost-admin.sharepoint.com" -replace " " , ""
                        Connect-SPOService -Url $UrlSharepoint -credential $O365Cred -ErrorVariable errordescr -ErrorAction SilentlyContinue
                        $CurrentError = $errordescr.message
                        #Credentials check
                        &$Global:CredentialValidation
                        }


#endregion Common Script Blocks


Function Request-Credential { 
    # Request for new credentials function. In the case of bad username or password or if the credentials needs to be different (connecting to another tenant).
     &$Global:UserCredential
}

Function Connect-O365PS { # Function to connecto to O365 services

    # Parameter request and validation
    param ([ValidateSet("Msol","AzureAd","AzureAdPreview","Exo","ExoBasic","Exo2","Eop","Scc","AIPService","Spo","Sfb","Teams")][Parameter(Mandatory=$true)] 
            $O365Service 
    )
    $Try = 0
    $global:errordesc = $null
    $Global:O365Cred=$null
    . $script:modulePath\ActionPlans\Connect-ServicesWithTokens.ps1 

    
#region Module Checks

    # Azure Module is mandatory

    if (!($global:addTypeAzureAD))
    {
        $minimumVersionAzureAD = '2.0.2.16'
        $CurrentProperty = "Checking AzureAD Module"
        if ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD}).count -gt 0)
        {
            $pathModule = split-path ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD } | Sort-Object -Property version -Descending)[0]).Path -parent
            
            $path = join-path $pathModule 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
            try 
            {
                Add-Type -Path $path
            }
            catch {}
            
            $CurrentDescription = "Azure AD Module for Windows PowerShell is installed and version is $(Split-Path $pathModule -Leaf)"
            $global:addTypeAzureAD = $true
        
        }
        else 
        {
            $CurrentDescription = "Azure AD Module for Windows PowerShell is not installed or version is less than $minimumVersionAzureAD. Initiated install from PowerShell Gallery"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            #Uninstall-Module AzureAD -Force -Confirm:$false -ErrorAction SilentlyContinue |Out-Null
            
            Install-Module AzureAD -Force -Confirm:$false -AllowClobber
            $pathModule = split-path ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD } | Sort-Object -Property version -Descending)[0]).Path -parent
            $path = join-path $pathModule 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
            try 
            {
                Add-Type -Path $path -IgnoreWarnings:$true -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            }
            catch {}
            
            $CurrentDescription = "Azure AD Module for Windows PowerShell is installed and version is $(Split-Path $pathModule -Leaf)"
            $global:addTypeAzureAD = $true
        }
    }

    # Checking if required modules are installed
    If ( $O365Service -eq "MSOL") {
        $updateMSOL = $false
        [version]$minimumVersion = "1.0.8070" 

        If ((get-module -ListAvailable -Name MSOnline).count -eq 0 )
        {
            $updateMSOL = $true
        }
        else 
        {
            $updateMSOL = $true
            foreach ($version in (get-module -ListAvailable -Name MSOnline).Version)
            {
                if ($version -ge $minimumVersion)
                {
                    $updateMSOL = $false
                }
            }
        }
        if ($updateMSOL)
        {
            $CurrentProperty = "Checking MSOL Module"
            Write-Host "`nMSOL Module for Windows PowerShell is not installed. Initiated install from PowerShell Gallery" -ForegroundColor Red
            $CurrentDescription = "MSOL Module for Windows PowerShell is not installed or is less than required version $minimumVersion. Initiated install from PowerShell Gallery"
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Uninstall-Module MSOnline -Force -Confirm:$false -ErrorAction SilentlyContinue |Out-Null
            Install-Module MSOnline -Force -Confirm:$false -AllowClobber
        }
    }

    <# Removed as AzureAD is mandatory for all connections to request manually the token
    If ( $O365Service -eq "AzureAD") {
        $updateAzureAD = $false
        [version]$minimumVersion = "2.0.0.131"

        If ((get-module -ListAvailable -Name AzureAD).count -eq 0 )
        {
            $updateAzureAD = $True
        }
        else 
        {
            $updateAzureAD = $True
            foreach ($version in (get-module -ListAvailable -Name AzureAD).Version)
            {
                if ($version -ge $minimumVersion)
                {
                    $updateAzureAD = $false
                }
            }
        }
        if ($updateAzureAD)
        {
            $CurrentProperty = "Checking AzureAD Module"
            $CurrentDescription = "Azure AD Module for Windows PowerShell is not installed or version is less than $minimumVersion. Initiated install from PowerShell Gallery"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            #Uninstall-Module AzureAD -Force -Confirm:$false -ErrorAction SilentlyContinue |Out-Null
            Install-Module AzureAD -Force -Confirm:$false -AllowClobber
        }
    }
    #>

    #TODO: need to check if AzureADPreview for sign-in logs 
    If ( $O365Service -eq "AzureADPreview") {
        $updateAzureADPreview = $false
        [version]$minimumVersion = "2.0.2.89"

        If ((get-module -ListAvailable -Name AzureADPreview).count -eq 0 )
        {
            $updateAzureADPreview = $True
        }
        else 
        {
            $updateAzureADPreview = $True
            foreach ($version in (get-module -ListAvailable -Name AzureADPreview).Version)
            {
                if ($version -ge $minimumVersion)
                {
                    $updateAzureADPreview = $false
                }
            }
        }
        if ($updateAzureADPreview)
        {
            $CurrentProperty = "Checking AzureADPreview Module"
            $CurrentDescription = "AzureADPreview Module is not installed or version is less than $minimumVersion. Initiated install from PowerShell Gallery"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            #Uninstall-Module AzureADPreview -Force -Confirm:$false -ErrorAction SilentlyContinue |Out-Null
            Install-Module AzureADPreview -Force -Confirm:$false -AllowClobber
        }
    }

    If ( $O365Service -eq "AIPService") {
        If ((Get-Module -ListAvailable -Name AIPService).count -eq 0) {
            $CurrentProperty = "Checking AIPService Module"
            $CurrentDescription = "AIPService needs to be updated or you have just updated without restarting the PC/laptop. Script will install the AIPService module from PowerShel Gallery"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Install-Module -Name AIPService -Force -Confirm:$false -AllowClobber
            Write-Host "Installed the AIPService module"
            #TODO: if AADRM module was installed, that needs to be uninstalled
            #TODO: check if AADRM Module was succesfully installed
        }
    }
    
    # Currently disabled as connection is done based on access token
    <# 
    if ( $O365Service -eq "Exo") {
        if ($null -eq ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | `
            Where-Object { $_ -notmatch "_none_" })) 
        {
            Write-Host "You requested to connect to Exchange Online with MFA but you don't have Exchange Online Remote PowerShell Module installed" -ForegroundColor Red
            Write-Host "Please check the article https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps" -ForegroundColor Red
            Write-Host  "The script will now exit" -ForegroundColor Red
            Read-Key
            Write-Log  -function Connect-O365PS -step "Connect to EXO with Modern & MFA" -Description "Required module is not installed."
            Disconnect-All
            Exit
        }
    }
    #>

    if ( $O365Service -eq "Exo2") {
        if ((Get-Module -ListAvailable -Name ExchangeOnlineManagement).count -eq 0) 
        {
            $CurrentProperty = "Checking ExchangeOnlineManagement v2 Module"
            $CurrentDescription = "ExchangeOnlineManagement module is not installed. We'll install it to support connectin to Exchange Online Module v2"
            write-host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Install-Module -Name ExchangeOnlineManagement -Force -Confirm:$false -AllowClobber
        }
    }

    # TODO: SPO prerequisites & modern module check

    # TODO: SFB prerequisites & modern module check
   
    #$Global:proxy = Read-Host
    if ($null -eq $Global:proxy) {
        Write-Host "`nAre you able to access Internet from this location without a Proxy?" -ForegroundColor Cyan
        $Global:proxy = get-choice "Yes", "No"
        $Global:PSsettings = New-PSSessionOption -SkipRevocationCheck 
        if ($Global:proxy -eq "n") {
            $Global:PSsettings = New-PSSessionOption -ProxyAccessType IEConfig -SkipRevocationCheck 
            
            if ($PSVersionTable.PSVersion.Major -eq 7) {
                (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer

                #Write-Host "Please input proxy server address (e.g.: http://proxy): " -ForegroundColor Cyan -NoNewline
                #$proxyServer = Read-Host
                #Write-Host "Please input proxy server port: " -ForegroundColor Cyan -NoNewline
                #$proxyPort = Read-Host
                $proxyConnection = "http://"+(Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer

            }

            else {
                #It doesn't work in PowerShell7
                $proxyConnection = ([System.Net.WebProxy]::GetDefaultProxy()).Address.ToString()
            }
            Invoke-WebRequest -Proxy $proxyConnection  -ProxyUseDefaultCredentials https://provisioningapi.microsoftonline.com/provisioningwebservice.svc
        }
    
    }

#endregion Module Checks

#region Connection scripts region
  if ($global:MfaOption -eq 0)
  {
      do 
      {

        Write-Host "Do your account require MFA to authenticate? (y/n): " -ForegroundColor Cyan -NoNewline
        $mfa =Read-Host 
        $mfa = $mfa.ToLower()
        if ($mfa -eq "y")
        {
            $global:MfaOption = 1
            Write-Host $global:MfaDisclaimer -ForegroundColor Red 
            $global:userPrincipalName = Get-ValidEmailAddress("UserPrincipalName used to connect to Office 365 Services")
            }
        if ($mfa -eq "n")
        {
            $global:MfaOption = 2
            $global:credentials = Get-Credential -Message "Please input your Office 365 credentials:"
            $global:userPrincipalName =$global:credentials.UserName
        }
      } until (($mfa -eq "y") -or ($mfa -eq "n"))
    

  }
  switch ($O365Service) {
    # Connect to MSOL
    "MSOL" {
        # Defining the banner variable and clear the errors
        $Global:Error.Clear();
        $Global:banner = "MSOL PowerShell"
        $CurrentProperty = "Connect MSOL"
        
        if ($global:MfaOption -eq 2)
        {
            if (!(Get-Module MSOnline)) {
                Import-Module MSOnline -Global -DisableNameChecking  -ErrorAction SilentlyContinue | Out-Null
            }
            $token = Get-TokenFromCache("AzureGraph")
            if ($null -eq $token)
            {
                $token = Get-Token("AzureGraph")
                Connect-MsolService -AdGraphAccessToken $token.AccessToken
            }
            else 
            {
                try {
                    $null = Get-MsolCompanyInformation -ErrorAction Stop
                }
                catch 
                {
                    Connect-MsolService -AdGraphAccessToken $token.AccessToken
                }
            }
        }
        elseif ($global:MfaOption -eq 1) {
            Do {

                    $errordescr = $null
                    $try++
                    try 
                    {
                        $null = Get-MsolCompanyInformation -ErrorAction Stop
                    }
                    catch 
                    {
                        Write-Host "$CurrentProperty"
                        if (!("MSOnline" -in (Get-Module).name))
                        {
                            Import-Module MSOnline -Global -DisableNameChecking  -ErrorAction SilentlyContinue | Out-Null
                        }
                        $errordescr = $null
                        Connect-MsolService -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        if ($null -eq $Global:Domain)
                        {
                            $Global:Domain = (get-msoldomain -ErrorAction SilentlyContinue -ErrorVariable errordescr| Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"}).Name
                        }
                        $CurrentError = $errordescr.exception.message
                    }
                    # Creating the session for PS MSOL Service
                    &$Global:CredentialValidation
            } while (($Try -le 2) -and ($null -ne $errordescr))
        }   
        &$Global:DisplayConnect
    }

    "AzureAD" {
        # Defining the banner variable and clear the errors
        $Global:Error.Clear();
        $Global:banner = "AzureAD PowerShell"
        $CurrentProperty = "Connect Azure"
        if ($global:MfaOption -eq 2)
        {
            if (!(Get-Module AzureAD)) {
                Import-Module AzureAD -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
            }
            $token = Get-TokenFromCache("AzureGraph")
            if ($null -eq $token)
            {
                $token = Get-Token("AzureGraph")
                Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue |Out-Null
            }
            else 
            {
                try {
                    $null = Get-AzureADTenantDetail -ErrorAction Stop
                }
                catch 
                {
                    Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue |Out-Null
                }
            }
        }
        elseif ($global:MfaOption -eq 1) {
        Do {
                # Defining the banner variable and clear the errors
                $Global:Error.Clear();
                $Global:banner = "AzureAD PowerShell"
                $errordescr = $null
                $try++
                try 
                {
                    $null = Get-AzureADTenantDetail -ErrorAction Stop
                }
                catch 
                {
                    Write-Host "$CurrentProperty"
                    if (!("AzureAD" -in (Get-Module).name))
                    {
                        Import-Module AzureAD -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                    }
                    $errordescr = $null
                    Connect-AzureAd -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                    if ($null -eq $Global:Domain)
                    {
                        $Global:Domain = (Get-AzureADDomain -ErrorAction SilentlyContinue -ErrorVariable errordescr| Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"}).Name
                    }
                    $CurrentError = $errordescr.exception.message
                }
                # Creating the session for PS MSOL Service
                &$Global:CredentialValidation
        }
        while (($Try -le 2) -and ($null -ne $errordescr))
    }
        &$Global:DisplayConnect
    }

    "AzureADPreview" {
        # Defining the banner variable and clear the errors
        $Global:Error.Clear();
        $Global:banner = "AzureADPreview PowerShell"
        $CurrentProperty = "Connect AzureADPreview"
        
        # Temporary: Forcing to connect to AzureAD only with promts
        # With AADGraph token Get-AzureADAuditSignInLogs fails with: Object reference not set to an instance of an object
        # With both AADGraph and MSGraph is failing with: 
        # User missing required MsGraph permission to access this API, please get any of the following permission for the user: AuditLog.Read.All
        #$global:MfaOption
        # 1 - Modern with MFA
        # 2 - Modern with no MFA (not checking to avoid connection based on AADGraph - checked against 3 should happen)
        if ($global:MfaOption -eq 3)
        {
            if (!(Get-Module AzureADPreview)) {
                Import-Module AzureADPreview -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
            }
            $token = Get-TokenFromCache("AzureGraph")
            if ($null -eq $token)
            {
                $token = Get-Token("AzureGraph")
                AzureADPreview\Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue |Out-Null
            }
            else 
            {
                try {
                    $null = AzureADPreview\Get-AzureADTenantDetail -ErrorAction Stop
                }
                catch 
                {
                    AzureADPreview\Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue |Out-Null
                }
            }
        }
        elseif (($global:MfaOption -eq 1) -or ($global:MfaOption -eq 2)) {
            Do {
                $errordescr = $null
                $try++
                try 
                {
                    $null = AzureADPreview\Get-AzureADTenantDetail -ErrorAction Stop
                }
                catch 
                {
                    Write-Host "$CurrentProperty"
                    if (!("AzureADPreview" -in (Get-Module).name))
                    {
                        Import-Module AzureADPreview -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
                    }
                    $errordescr = $null
                    AzureADPreview\Connect-AzureAD -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                    if ($null -eq $Global:Domain)
                    {
                        $Global:Domain = (AzureADPreview\Get-AzureADDomain -ErrorAction SilentlyContinue -ErrorVariable errordescr| Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"}).Name
                    }
                    $CurrentError = $errordescr.exception.message
                }
                # Creating the session for PS MSOL Service
                &$Global:CredentialValidation
            } while (($Try -le 2) -and ($null -ne $errordescr))
        }

        &$Global:DisplayConnect
    }
    # Connect to Exchange Online PowerShell
    "EXO" {    
        # Defining the banner variable and clear the errors
        $CurrentProperty = "Connect EXO"
        $Global:Error.Clear();
        $Global:banner = "Exchange Online PowerShell - Modern"

        if ($global:MfaOption -eq 2)
        {
            $token = Get-TokenFromCache("EXO")
            if ($null -eq $token)
            {
                $token = Get-Token("EXO")
                Get-PSSession -name EXO -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
                
                # Build the auth information
                $Authorization = "Bearer {0}" -f $Token.AccessToken
                $UserId = ($Token.UserInfo.DisplayableId).tostring()
                
                # create the "basic" token to send to O365 EXO
                $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
                $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
                
                # Create and import the session
                $Session = New-PSSession -Name EXO -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
                Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue'
            }
            else 
            {
                try {
                    $null = Get-OrganizationConfig -ErrorAction Stop
                }
                catch 
                {
                    Get-PSSession -name EXO -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
                
                    # Build the auth information
                    $Authorization = "Bearer {0}" -f $Token.AccessToken
                    $UserId = ($Token.UserInfo.DisplayableId).tostring()
                    
                    # create the "basic" token to send to O365 EXO
                    $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
                    $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
                    
                    # Create and import the session
                    $Session = New-PSSession -Name EXO -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
                    Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue'
                }
            }
        }
        elseif ($global:MfaOption -eq 1)
        {
             # The loop for re-entering credentials in case they are wrong and for re-connecting
            Do {
                

                $try++

                try 
                {
                    $null = Get-OrganizationConfig -ErrorAction Stop
                }
                catch 
                {
                        Write-Host "$CurrentProperty"
                        if (!("Microsoft.Exchange.Management.ExoPowershellModule" -in (Get-Module).Name))
                        {
                            Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | `
                                Where-Object { $_ -notmatch "_none_" } | Select-Object -First 1) -Global -DisableNameChecking -Force -ErrorAction SilentlyContinue
                        }
        
                        $errordescr = $null
                        if (($null -eq $Global:EXOSession )-or ($Global:EXOSession.State -eq "Closed") -or ($Global:EXOSession.State -eq "Broken"))
                        {
                            $Global:EXOSession = New-ExoPSSession -UserPrincipalName $global:UserPrincipalName -PSSessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop
                            $CurrentError = $errordescr.exception 
                            Import-Module (Import-PSSession $EXOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue
                            $null = Get-OrganizationConfig -ErrorAction SilentlyContinue -ErrorVariable errordescr
                            $CurrentError = $errordescr.exception.message + $Global:Error[0]
                        }

                }
                &$Global:CredentialValidation
            } while (($Try -le 2) -and ($Global:Error)) 
        }
        &$Global:DisplayConnect
    }

    # Connecto to EXO Basic Authentication (not recommended unless you want to test secifically BasicAuth)
    "ExoBasic" {
        If ($null -eq $Global:O365Cred) {
            &$Global:UserCredential
        }
        try {

        $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $global:O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop 
        $CurrentError = $errordescr.exception  
        Import-Module (Import-PSSession $EXOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue
        $CurrentDescription = "Success"
        $Global:Domain = Get-AcceptedDomain | Where-Object{$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"}  

        }
        catch 
        {
            $CurrentDescription = "`""+$CurrentError.ErrorRecord.Exception +"`""
        } 
        &$Global:DisplayConnect
    }
    # Connect to EXO2
    "EXO2" {

        $CurrentProperty = "Connect EXOv2"
        Do 
        {
            # Defining the banner variable and clear the errors
            $Global:Error.Clear();
            $Global:banner = "EXOv2 PowerShell"
            $errordescr = $null
            $try++
            try 
            {
                $null = Get-EXOMailbox -ErrorAction Stop
            }
            catch 
            {
                Write-Host "$CurrentProperty"
                if (!("ExchangeOnlineManagement" -in (Get-Module).name))
                {
                    Import-Module ExchangeOnlineManagement -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                }
                $errordescr = $null
                Connect-ExchangeOnline  -PSSessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue -ShowBanner:$false
                $null = get-EXOMailbox -ErrorVariable errordescr
                $CurrentError = $errordescr.exception.message
            }
            # Creating the session for PS MSOL Service
            &$Global:CredentialValidation
            }
            while (($Try -le 2) -and ($null -ne $errordescr))
            &$Global:DisplayConnect
    }

    # Connect to EOP
    "EOP" {
       # Defining the banner variable and clear the errors
       $CurrentProperty = "Connect EOP"
       $Global:Error.Clear();
       $Global:banner = "Exchange Online Protection PowerShell - Modern"

       if ($global:MfaOption -eq 2)
       {
           $token = Get-TokenFromCache("EXO")
           if ($null -eq $token)
           {
               $token = Get-Token("EXO")
               Get-PSSession -name EOP -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
               
               # Build the auth information
               $Authorization = "Bearer {0}" -f $Token.AccessToken
               $UserId = ($Token.UserInfo.DisplayableId).tostring()
               
               # create the "basic" token to send to O365 EXO
               $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
               $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
               
               # Create and import the session
               $Session = New-PSSession -Name EOP -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.protection.outlook.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
               Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue'
           }
           else 
           {
               try {
                   $null = Get-OrganizationConfig -ErrorAction Stop
               }
               catch 
               {
                   Get-PSSession -name EOP -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
               
                   # Build the auth information
                   $Authorization = "Bearer {0}" -f $Token.AccessToken
                   $UserId = ($Token.UserInfo.DisplayableId).tostring()
                   
                   # create the "basic" token to send to O365 EXO
                   $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
                   $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
                   
                   # Create and import the session
                   $Session = New-PSSession -Name EOP -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.protection.outlook.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
                   Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue'
               }
           }
       }
       elseif ($global:MfaOption -eq 1)
       {
            # The loop for re-entering credentials in case they are wrong and for re-connecting
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Exchange Online Protection PowerShell"
                        $try++
                        # Creating EOP PS session
                        $Global:EOPSession = New-PSSession -ConfigurationName EOP -ConnectionUri "https://ps.protection.outlook.com/powershell-liveid" -Credential $global:O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $EOPSession  -AllowClobber -DisableNameChecking ) -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                        $Global:Domain = Get-AcceptedDomain | Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"} 
            }
            while (($Try -le 2) -and ($Global:Error)) 
        }
        &$Global:DisplayConnect
    }

    # Connect to Compliance Center Online
    "SCC" {
    # The loop for re-entering credentials in case they are wrong and for re-connecting
    # Defining the banner variable and clear the errors
    $CurrentProperty = "Connect SCC"
    $Global:Error.Clear();
    $Global:banner = "Security & Compliance Online PowerShell - Modern"

    if ($global:MfaOption -eq 2)
    {
        $token = Get-TokenFromCache("EXO")
        if ($null -eq $token)
        {
            $token = Get-Token("EXO")
            Get-PSSession -name SCC -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
            
            # Build the auth information
            $Authorization = "Bearer {0}" -f $Token.AccessToken
            $UserId = ($Token.UserInfo.DisplayableId).tostring()
            
            # create the "basic" token to send to O365 EXO
            $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
            $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
            
            # Create and import the session
            $Session = New-PSSession -Name SCC -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
            Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue' -Prefix cc
        }
        else 
        {
            try {
                $null = Get-OrganizationConfig -ErrorAction Stop
            }
            catch 
            {
                Get-PSSession -name SCC -ErrorAction SilentlyContinue | Remove-PSSession -Confirm:$false
            
                # Build the auth information
                $Authorization = "Bearer {0}" -f $Token.AccessToken
                $UserId = ($Token.UserInfo.DisplayableId).tostring()
                
                # create the "basic" token to send to O365 EXO
                $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
                $Credtoken = New-Object System.Management.Automation.PSCredential($UserId, $Password)
                
                # Create and import the session
                $Session = New-PSSession -Name SCC -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true' -Credential $Credtoken -SessionOption $Global:PSsettings -Authentication Basic -AllowRedirection -ErrorAction Stop
                Import-Module (Import-PSSession $Session -AllowClobber) -Global -WarningAction 'SilentlyContinue'
            }
        }
    }
    elseif ($global:MfaOption -eq 1)
    {
        Do {
            # Defining the banner variable and clear the errors
            $Global:Error.Clear();
            $Global:banner = "Security&Compliance Online PowerShell - Modern & MFA"
            $try++

            try 
            {
                $null = Get-ccLabel -ErrorAction Stop
            }
            catch 
            {
                Write-Host "$CurrentProperty"
                if (!("Microsoft.Exchange.Management.ExoPowershellModule" -in (Get-Module).Name))
                {
                    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | `
                        Where-Object { $_ -notmatch "_none_" } | Select-Object -First 1) -Global -DisableNameChecking -Force -ErrorAction SilentlyContinue
                }

                $errordescr = $null
                if (($null -eq $Global:IPPSession )-or ($Global:IPPSession.State -eq "Closed") -or ($Global:IPPSession.State -eq "Broken"))
                {
                    $Global:IPPSession = New-ExoPSSession -UserPrincipalName $global:UserPrincipalName -ConnectionUri 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId' `
                    -PSSessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop
                    $CurrentError = $errordescr.exception 
                    Import-Module (Import-PSSession $IPPSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue -Prefix cc
                    $null = Get-ccLabel -ErrorAction SilentlyContinue -ErrorVariable errordescr
                    $CurrentError = $errordescr.exception.message + $Global:Error[0]
                }
            }
            &$Global:CredentialValidation
        } while (($Try -le 2) -and ($Global:Error)) 
    }
        &$Global:DisplayConnect
    }
    
    #Connect to SharePoint Online PowerShell
    "SPO" {
        $Global:Error.Clear();
        Import-Module MSOnline ;
        If ($null -eq $Global:O365Cred) {
                &$Global:UserCredential
        }
        # The loop for re-entering credentials in case they are wrong and for re-connecting
        $CurrentProperty = "Connect SPO"
        
        Do {
                #### Update_Razvan (conditie pentru admin care nu foloseste onmicrosoft.com si verificare conectare MSOL pentru a lua domeniul)
            If ($O365Cred.UserName -like "*.onmicrosoft.com") {
                If ($O365Cred.UserName -like "*.mail.onmicrosoft.com") {
                    $DomainHost = (($O365Cred.UserName -split ".mail.onmicrosoft.com")[0].Substring(0) -split "@")[1].Substring(0)
                    &$Global:SPOConnectBlock
                }
            $DomainHost = (($O365Cred.UserName -split ".onmicrosoft.com")[0].Substring(0) -split "@")[1].Substring(0)
            &$Global:SPOConnectBlock
            }

            Else {
                If ($null -ne $domain) {
                    # Substract the domain host name out of the tenant name
                    $DomainHost = ($domain.name -split ".onmicrosoft.com") 
                    &$Global:SPOConnectBlock
                }
                Else {
                    $Global:Error.Clear();
                    $Global:banner = "SharePoint Online PowerShell"
                    $try++
                    $URL= read-host "Please Input the connection URL (i.e.: https://Tenant_Domain-admin.sharepoint.com/)"
                    Connect-SPOService -Url $URL -credential $O365Cred -ErrorVariable errordescr -ErrorAction SilentlyContinue
                    &$Global:CredentialValidation
                }
            }
        }
        while (($Try -le 2) -and ($null -ne $Global:Error)) 
        &$Global:DisplayConnect                     
    }

    # Connect to Skype Online PowerShell
    "SFB" {
                $Global:Error.Clear();
                Import-Module MSOnline ;
                If ($null -eq $Global:O365Cred) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                $CurrentProperty = "Connect SFB"
                
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Skype for Business Online PowerShell"
                        $try++
                        # Import SFB Online PS module
                        Import-Module LyncOnlineConnector
                        # Creating a new SFB Online PS Session
                        $global:sfboSession = New-CsOnlineSession -Credential $global:O365Cred -ErrorVariable errordescr
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $sfboSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking 
                        # Credentials check
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($null -ne $Global:Error)) 
                &$Global:DisplayConnect
    }

    # Connect to AIPService PowerShell
    "AIPService" {
        # Defining the banner variable and clear the errors
        $Global:Error.Clear();
        $Global:banner = "AIP PowerShell"
        $CurrentProperty = "Connect AIP"
        
        if ($global:MfaOption -eq 2)
        {
            if (!(Get-Module AIPService)) {
                Import-Module AIPService -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
            }
            $token = Get-TokenFromCache("AIPService")
            if ($null -eq $token)
            {
                $token = Get-Token("AIPService")
                Connect-AipService -AccessToken $token.AccessToken -ErrorAction SilentlyContinue | Out-Null
            }
            else 
            {
                try {
                    $null = Get-AipServiceConfiguration -ErrorAction Stop
                }
                catch 
                {
                    Connect-AipService -AccessToken $token.AccessToken -ErrorAction SilentlyContinue | Out-Null
                }
            }
        }
        elseif ($global:MfaOption -eq 1) {
            do
            {
                $Global:Error.Clear();
                $Global:banner = "AIPService PowerShell"
                $errordescr = $null
                $try++
                try
                {
                    $null = Get-AipServiceConfiguration -ErrorAction Stop
                }
                catch 
                {
                    Write-Host "$CurrentProperty"
                    if (!("AIPService" -in (Get-Module).name))
                    {
                        Import-Module AIPService -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                    }
                    $errordescr = $null
                    $Global:Error.Clear();
                    Connect-AipService -ErrorVariable errordescr -ErrorAction SilentlyContinue
                    $null = Get-AipServiceConfiguration -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                    $CurrentError = $errordescr.exception.message
                }
                &$Global:CredentialValidation
            } while (($Try -le 2) -and ($null -ne $Global:Error)) 
        }
        &$Global:DisplayConnect
    }
  }
}
#endregion Connection scripts region


Function Set-GlobalVariables {
    Clear-Host
    Write-Host 
    $global:FormatEnumerationLimit = -1
    $script:PSModule = $ExecutionContext.SessionState.Module
    $script:modulePath = $script:PSModule.ModuleBase
    $global:ts = Get-Date -Format yyyyMMdd_HHmmss
    $global:Path =[Environment]::GetFolderPath("Desktop")
    $Global:Path += "\PowerShellOutputs"
    $global:WSPath = "$Path\PowerShellOutputs_$ts"
    $global:starline = New-Object String '*',5
    $global:addTypeAzureAD = $false
    $global:MfaOption = 0;
    #$Global:ExtractXML_XML = "Get-MigrationUserStatistics ", "Get-ImapSubscription "
    $global:Disclaimer ='Note: Before you run the script: 

    The sample scripts are not supported under any Microsoft standard support program or service. 
    The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims 
    all implied warranties including, without limitation, any implied warranties of merchantability 
    or of fitness for a particular purpose. The entire risk arising out of the use or performance of 
    the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
    or anyone else involved in the creation, production, or delivery of the scripts be liable for any 
    damages whatsoever (including, without limitation, damages for loss of business profits, business 
    interruption, loss of business information, or other pecuniary loss) arising out of the use of or 
    inability to use the sample scripts or documentation, even if Microsoft has been advised of the 
    possibility of such damages.
    '
    $global:MfaDisclaimer = @"
    1. Please note that not all services allow PowerShell connection with MFA! Example: AIPService module doesn`'t support MFA
    2. Using an account with MFA will generate multiple prompts for different Office 365 workloads
    3. If you require to have an account with MFA, you can set trusted IPs from which MFA prompt won't be required. See article: https://docs.microsoft.com/azure/active-directory/authentication/howto-mfa-mfasettings#trusted-ips
"@

    Write-Host $global:Disclaimer -ForegroundColor Red
    Start-Sleep -Seconds 3
    
    if (!(Test-Path $Path)) 
    {
        Write-Host "We are creating the following folder $Path"
        New-Item -Path $Path -ItemType Directory -Confirm:$False |Out-Null
    }

    if (!(Test-Path $WSPath))
    {
    Write-Host "We are creating the following folder $WSPath"
    New-Item -Path $WSPath -ItemType Directory -Confirm:$False |Out-Null
    }
    
    $global:outputFile = "$WSPath\Log_$ts.csv"
    $global:columnLabels = "Time, Function, Step, Description"
    Out-File -FilePath $outputFile -InputObject $columnLabels -Encoding UTF8 |Out-Null

    Set-Location $WSPath
    Write-Host "`n"

    #if ($null -eq $global:credential)
    #{
        #$global:userPrincipalName = Get-ValidEmailAddress("UserPrincipalName used to connect to Office 365 Services")
        #Write-Host "Please note that depening the Office 365 Services we need to connect, you might be asked to re-add the UserPrincipalName in another Authentication Form!" -ForegroundColor Yellow
        #Start-Sleep -Seconds 5
    #}

    
}

function Get-ValidEmailAddress([string]$EmailAddressType)
{
    [int]$count = 0
    do
    {
        Write-Host "Enter Valid $EmailAddressType`: " -ForegroundColor Cyan -NoNewline
        [string]$EmailAddress = Read-Host
        [bool]$valid = ($EmailAddress.Trim() -match "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$")
        $count++
    }
    while (!$valid -and ($count -le 2))
    
    if ($valid)
    {
        return $EmailAddress.Trim()
    }
    else 
    {   
        [string]$Description = "Received 3 invalid email address inputs, the script will return to O365Troubleshooters Main Menu"
        Write-Host "`n$Description" -ForegroundColor Red
        Start-Sleep -Seconds 3
        Write-Log -function "Get-ValidEmailAddress" -step "input address" -Description $Description
        Read-Key
        Start-O365TroubleshootersMenu
    }   
    
<# Old recurse of the function - replaced by new version with counter.
    if($EmailAddress -match "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$")
    {
        
    }
    else
    {
        Get-ValidEmailAddress($EmailAddressType)
    }
#>
}

function Read-IntFromConsole {
    param ([string][Parameter(Mandatory=$true)]$IntType)

    [int]$count = 0

    do
    {
        [bool]$valid = $true
        Write-Host "Enter Valid $IntType`: " -ForegroundColor Cyan -NoNewline
        try
        {
            [int]$IntFromConsole = Read-Host
            if($IntFromConsole -eq 0)
            {
                throw "System.Management.Automation.RuntimeException"
            }
        }
        catch [System.Management.Automation.RuntimeException]
        {
            Write-Host "Invalid $IntType returned" -ForegroundColor Red
            $valid = $false
            $count++
        }
    }while (!$valid -and ($count -le 2))
    
    if ($valid)
    {
        return $IntFromConsole
    }
    else 
    {   
        [string]$Description = "Received 3 invalid $IntType inputs, the script will return to O365Troubleshooters Main Menu"
        Write-Host "`n$Description" -ForegroundColor Red
        Start-Sleep -Seconds 3
        Write-Log -function "Read-IntFromConsole" -step "input number" -Description $Description
        Read-Key
        Start-O365TroubleshootersMenu
    }
}
Function New-XMLObject {
    param ( 
        $CmdletsNeeded
        )

 
    
        
$Global:Error.Clear()

$InitialErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Stop"
        
        #Log-Write -ScriptName 
        write-log -Function "New-XMLObject" -step "Started"
            
    $Global:Obj = New-Object Object | Select-Object -Property $CmdletsNeeded 
        
    #$Obj

    $Obj.psobject.Properties | 
        ForEach-Object {

        try {
    
            $Global:Error.Clear()
            $CurrentProperty = $_.name
            $Obj.$CurrentProperty = Invoke-Expression $CurrentProperty
            $CurrentDescription = $_
            write-host "Function New-XMLObject, current step: $CurrentProperty"
            
            
            }
      catch {
            
            $myerror = $Global:error[0]
            $CurrentDescription = $myerror.Exception.Message
            write-host $currentDescrioption
              
            $CurrentDescription = "`""+$CurrentDescription+"`"" 
                
            }
            write-log -Function "New-XMLObject" -Step $CurrentProperty -Description $CurrentDescription
        
            $myerror=$null
       }
    $ErrorActionPreference = $InitialErrorActionPreference
    return $Obj
    
}


function Write-Log {
    <# Example to use/call this function
        write-log -Function "Function Missing-Mailbox" -Step "Get_CASMailbox" -Description "Mailbox not found"
        write-log -Function "Function Missing-Mailbox" -Step "Get_CASMailbox" -Description $error[0]
    #>  
	param ($function, $step, $Description)
Write-Host

$tserror = Get-Date -Format yyyyMMdd_hhmmss
$currentRecord = "$tserror,$function,$step,$Description"
Out-File -FilePath $outputFile -InputObject $currentRecord -Encoding UTF8 -Append 
}


Function Get-Choice {
    <#
    Example:
    $Options = "White", "black"
    Get-Option $options
    $opt=$options[$option]
    write-host "You opted for $opt"
    #>
    
    param ( 
        $OptionsList
        )
    
    [int]$i=0
    do
    {
        
        $OptionsList | ForEach-Object  {
                write-host "Press $($_[0]) for '$_'"
        }
        [string]$Option=read-host "Please answer by typing first letter of the option" 
        [bool]$validAnswer=$false
        $OptionsList | ForEach-Object  {
            if ($_.ToLower()[0] -eq $Option.ToLower()[0]) 
            {
                $validAnswer = $true
            }
        }
        $i++
    }
     while (($validAnswer -eq $false) -and ($i -le 2))
     if ($validAnswer -eq $false)
     {
        Write-Log "Get-Choice" -step "provide one of the expected choices"    -Description "function received 3 consecutive unexpected answers, so will exit"
        Write-Host "You provided unexpected answers 3 consecutive times so the script will close" -ForegroundColor Red
        disconnect-all
        exit
     }
     return $Option
    
}


Function Open-URL {
        # Function for openining an URL
        param
        (
            [Parameter(Mandatory = $true, HelpMessage = 'URL to open')]
            $URL
        )
        Write-Host "We are trying to open the following URL: " -NoNewline
        Write-Host $URL -ForegroundColor "green"
        Start-Process -FilePath iexplore.exe -ArgumentList $URL
        return
    } 

Function Test-DotNet {
    # Function to check if .net 4.5 is installed
    $Global:Error.Clear()
    write-log -Function "Test-DotNet" -step "Started"
    
    $ndpDirectory = 'hklm:\SOFTWARE\Microsoft\NET Framework Setup\NDP\'
    $v4Directory = "$ndpDirectory\v4\Full"
    
    Write-Host "`nWe are trying to determine which .NET Framework version is installed"

    if (Test-Path $v4Directory)
        {
               $version = Get-ItemProperty $v4Directory -name Version | Select-Object -expand Version
               $dotnet=($version).Split(".")
               If (($dotnet[0] -eq 4) -and ($dotnet[1] -ge 5))
               {
               Write-Host "You have the following .NET Framework version installed: " $version
               Write-Host "The .NET Framework version meets the minimum requirements" -foregroundcolor "green"
               }
               else
               {
               Write-Host "You have the following .NET Framework version installed: " $version
               Write-Host "Your .net version is less than 4.5. Please update the .NET Framework version" -foregroundcolor "red"
               Open-URL ("http://go.microsoft.com/fwlink/?LinkId=671744")
               Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
               exit
               }
        }
        else
        {
            Write-Host "Your .net version is less than 4.5. Please update the .NET Framework version" -foregroundcolor "red"
            Open-URL ("http://go.microsoft.com/fwlink/?LinkId=671744")
            write-log -Function "Test-DotNet" -step "Check .net version" -Description "less than 4.5"
            Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
            exit
        }
        return
 }


Function Test-PSVers {
    # Function to check PS version
    $Global:Error.Clear()
    write-log -Function "Test-PSVers" -step "Started"
    Write-Host "`nWe are trying to determine the Operating System"
    # Display Operating System as it might help to identify what pack should be downloaded
    $op_ver = Get-WmiObject Win32_OperatingSystem | Select-Object Caption, OSArchitecture
    write-host "`nYou have the following Operating System:" $op_ver.Caption -foregroundcolor "green"
    write-host "You have the following Operating System Arhitecture:" $op_ver.OSArchitecture -foregroundcolor "green"
    $op_ver_srv = $op_ver.Caption.ToLower().Contains("Server".ToLower())
   
    # Check if the PoweShell version is 5
    Write-Host "`nWe are trying to determine which PowerShell version is installed"
    write-log -Function "Test-PSVers" -step "Check Powershell version" -Description $PSVersionTable.PSVersion.Major
    If ($PSVersionTable.PSVersion.Major -le 2)
    {
        If ($op_ver_srv)
         {
             Write-Host "`nYou have a Server Operating System !" -foregroundcolor "red"
             write-log -Function "Test-PSVers" -step "Check Powershell version" -Description $PSVersionTable.PSVersion.Major
             Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
             exit
            }
        else
        {
        Test-DotNet
        Write-Host "`nYou have the following Powershell version:" $PSVersionTable.PSVersion.Major
        Write-Host "`nYour Powershell version is less than 5. Please update your Powershell by installing Windows Management Framework 5.0 !" -foregroundcolor "magenta"
        Open-URL ("https://www.microsoft.com/en-us/download/details.aspx?id=50395")
        Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
        exit
        }
    }
     Else
    {
     Write-Host "`nYou have the following Powershell version:" $PSVersionTable.PSVersion.Major -foregroundcolor "green"
     }

}

function Start-Elevated {
    Write-Host "Starting new PowerShell Window with the O365Troubleshooters Module loaded"
    Read-Key
    Start-Process powershell.exe -ArgumentList "-noexit -Command Install-Module O365Troubleshooters -force; Import-Module O365Troubleshooters -force; Start-O365Troubleshooters -elevatedExecution `$true" -Verb RunAs #-Wait
    #Exit
}

function Disconnect-All {
    
    $CurrentDescription = "Disconnect is successful!"

    try {
            # Disconnect EXOv2
            if (("ExchangeOnlineManagement" -in (Get-Module).name))
            {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }

            # Disconnect AIP
            if (("AipService" -in (Get-Module).name))
            {
                Disconnect-AipService -Confirm:$false -ErrorAction SilentlyContinue
            }

            # Disconnect AzureAD
            if (("AzureAD" -in (Get-Module).name))
            {
                AzureAD\Disconnect-AzureAD -Confirm:$false -ErrorAction SilentlyContinue
                Remove-Module AzureAD -Force -Confirm:$false -ErrorAction SilentlyContinue
            }

            # Disconnect AzureADPreview
            if (("AzureADPreview" -in (Get-Module).name))
            {
                AzureADPreview\Disconnect-AzureAD -Confirm:$false -ErrorAction SilentlyContinue
                Remove-Module AzureADPreview -Force -Confirm:$false -ErrorAction SilentlyContinue
            }

            # Disconnect all PsSessions
            Get-PSSession | Remove-PSSession
    }

    catch {
             
             $CurrentDescription = "`""+$Global:Error[0].Exception.Message+"`"" 

    }

    write-log -Function "Disconnect - close sessions" -Step $CurrentProperty -Description $CurrentDescription

    
    $CurrentDescription = "Execution Policy was successfully set to its original value!"

    try
    {
            # Set back the initial ExecutionPolicy value
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy $_CurrentUser -Force -ErrorAction SilentlyContinue
    }
    catch 
    {
            $CurrentDescription = "`""+$Global:Error[0].Exception.Message+"`"" 
    }
    
    write-log -Function "Disconnect - ExecutionPolicy" -Step $CurrentProperty -Description $CurrentDescription
    # Read-Host -Prompt "Please press [Enter] to continue"
    Read-Key
    }

Function Read-Key{
    if ($psISE)
    {
        Write-Host "Press [Enter] to continue" -ForegroundColor Cyan
        Read-Host 
    }
    else 
    {
        Write-Host "Press any key to continue" -ForegroundColor Cyan
        #[void][System.Console]::ReadKey($true)
        $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")|Out-Null
    }

}


### <summary>
### Export-ReportToHTML function is used to convert any kind of report to HTML file
### </summary>
### <param name="FilePath">FilePath represents the Full path of the .html file that will be saved by this function </param>
### <param name="PageTitle">PageTitle represents the title of the Report. This will appear in the browser tab</param>
### <param name="ReportTitle">ReportTitle represents the title of the Report. This will appear inside the report, as a descriptive title of the report</param>
### <param name="TheObjectToConvertToHTML">
###     TheObjectToConvertToHTML represents the object that need to be converted to HTML
###         Its structure is:
###             [string]$SectionTitle - Contains the header of the data that will be added to the HTML file
###             [string]$SectionTitleColor - Contains the header of the data that will be added to the HTML file (accepted values are "Green" or "Red")
###             [string]$Description - Contains description of the data that will be added to the HTML file
###             [string]$DataType - Contains the data type of the data that need to be added into the HTML file (it can be: "Table", "String")
###             [ArrayList]/[String]$EffectiveData - Contains the effective data that need to be added into a HTML file
###             [String]$TableType - If EffectiveData is table, we need to know how to list it (accepted values: "List" = Vertical, "Table" = Horizontal)
### </param>
Function Export-ReportToHTML {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$false)]
        [string]$PageTitle,

        [Parameter(Mandatory=$false)]
        [string]$ReportTitle,

        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$TheObjectToConvertToHTML

    )

    ### Create header of the HTML file
    $header = @"
<!--HTML file created by O365Troubleshooters-->
<title>$PageTitle</title>
<style>
    h1 {
        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;
    }
    
    h2 {
        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;
    }

    h3 {
        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 12px;
    }

    .Green {
        color: #008000;
    }
    
    .Red {
        color: #ff0000;
    }

    .Black {
        color: #000000;
    }
        
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
    }
    
    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

    #CreationDate {
        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;
    }
</style>
<a href="https://www.powershellgallery.com/packages/O365Troubleshooters" target="_blank">
    <img src="https://raw.githubusercontent.com/vilega/O365Troubleshooters/master/Resources/O365Troubleshooters-Logo.png" alt="O365Troubleshooters" width="173" height="128">
</a>
"@

    [int]$i = 0
    [string]$TheBody = "<h1>$ReportTitle</h1>"
    
    ### For each scenario, convert the data to HTML
    foreach ($Entry in $TheObjectToConvertToHTML) {
        if ($Entry.DataType -eq "String") {
            $TheValue = ConvertTo-Html -PreContent "<h2 class=`"$($Entry.SectionTitleColor)`">`n`n$($Entry.SectionTitle)</h2><h3 class=`"Black`">`n$($Entry.Description)`n</h3>" -PostContent $($Entry.EffectiveData)
        }
        else {
            $TheValue = $($Entry.EffectiveData) | ConvertTo-Html -As $($Entry.TableType) -PreContent "<h2 class=`"$($Entry.SectionTitleColor)`">`n`n$($Entry.SectionTitle)</h2><h3 class=`"Black`">`n$($Entry.Description)`n</h3>"
        }

        ### Adding sections in the body of the HTML report
        $TheBody = $TheBody  + " " + $TheValue

        $i++
    }

    ### Create the report to convert the entire content to HTML
    $Report = ConvertTo-Html -Head $header -Body $TheBody -PreContent "<p>Creation Date: $((Get-date).ToUniversalTime()) UTC</p>"

    ### Export the HTML report to the specific location
    $Report | Out-File $FilePath
}


### <summary>
### Prepare-ObjectForHTMLReport function is used to prepare the objects to be converted to HTML file
### </summary>
### <param name="SectionTitle">SectionTitle represents the header of the section</param>
### <param name="SectionTitleColor">SectionTitleColor represents the color of the header of the section (valid values to use: "Black", "Green" or "Red")</param>
### <param name="Description">Description represents the description for the section</param>
### <param name="DataType">DataType represents the type of data for the section (valid values to use: "ArrayList" or "String")</param>
### <param name="EffectiveDataString">EffectiveDataString represents the effective data. This is the data used in case the DataType is "String"</param>
### <param name="EffectiveDataArrayList">EffectiveDataArrayList represents the effective data. This is the data used in case the DataType is "ArrayList"</param>
### <param name="TableType">TableType represents the type of table HTML should list.
###         This is available only if DataType is "ArrayList" (valid values to use: "List" or "Table")</param>
###
### <returns>TheObject - this is the object in which data that need to be converted to HTML is stored</returns>
function Prepare-ObjectForHTMLReport {
param (
    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [string]$SectionTitle,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [ValidateSet("Black", "Green", "Red")]
    [string]$SectionTitleColor,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [string]$Description,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [ValidateSet("ArrayList", "String")]
    [string]$DataType,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [string]$EffectiveDataString,

    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [PSCustomObject]$EffectiveDataArrayList,

    [Parameter(ParameterSetName = "ArrayList", Mandatory=$false)]
    [ValidateSet("List", "Table")]
    [string]$TableType
)

    ###Create the object, with all needed Properties, that will be used to convert into an HTML report
    $TheObject = New-Object PSObject
        $TheObject | Add-Member -NotePropertyName SectionTitle -NotePropertyValue $SectionTitle
        $TheObject | Add-Member -NotePropertyName SectionTitleColor -NotePropertyValue $SectionTitleColor
        $TheObject | Add-Member -NotePropertyName Description -NotePropertyValue $Description
        $TheObject | Add-Member -NotePropertyName DataType -NotePropertyValue $DataType
        if ($DataType -eq "ArrayList") {
            $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataArrayList
            $TheObject | Add-Member -NotePropertyName TableType -NotePropertyValue $TableType
        }
        else {
            $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataString
        }

    ### Return the created object
    return $TheObject

}


Function    Start-O365Troubleshooters
{
    param(
        [bool][Parameter(Mandatory=$false)] $elevatedExecution=$false
    )
    if (!$elevatedExecution)
    {
        Start-Elevated
    }
    else 
    {
        Set-GlobalVariables
        Start-O365TroubleshootersMenu
    }
}

Function Start-O365TroubleshootersMenu {
    $menu=@"
    1  Encryption: Office Message Encryption General Troubleshooting
    2  Mail Flow: SMTP Relay Test
    3  Migration: Analyze Mailbox move (Hybrid migration)
    4  Security: Compromised Tenant Investigation
    5  Tools: Exchange Online Audit Search
    6  Tools: Unified Logging Audit Search
    7  Tools: Azure AD Audit Sign In Log Search
    8  Tools: Find all users with a specific RBAC Role
    9  Tools: Find all users with all RBAC Roles
    10 Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox
    11 Tools: Decode SafeLinks URL
    12 Tools: Export Quarantine Messages
    13 Tools: Transform IMCEAEX (old LegacyExchangeDN) to X500 address
    14 Groups: DL to O365 Groups Upgrade Checker
    Q  Quit
     
    Select a task by number or Q to quit
"@
Clear-Host
Write-Host "Main Menu" -ForegroundColor Cyan
$r = Read-Host $menu

# Security: Analyze compromise account/tenant
# Write-Host "Action Plan: Analyze compromise account/tenant" -ForegroundColor Green
# . $script:modulePath\ActionPlans\Start-CompromisedInvestigation.ps1

Switch ($r) {
    "1" {
        Write-Host "Action Plan: Office Message Encryption General Troubleshooting" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-OfficeMessageEncryption.ps1
    }
     
    "2" {
        Write-Host "Action Plan: SMTP Relay Test" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-Office365Relay.ps1
    }

    "3" {
        Write-Host "Action Plan: Mailbox Migration - Hybrid" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-MailboxMigrationAnalyzer.ps1
    }

    "4" {
        Write-Host "Action Plan: Compromised Tenant" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-CompromisedInvestigation.ps1
        Start-CompromisedMain
    }

    "5" {
        Write-Host "Tools: Exchange Online Audit Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1
        Start-ExchangeOnlineAuditSearch
    }

    "6" {
        Write-Host "Tools: Unified Logging Audit Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-UnifiedAuditLogSearch.ps1
    }

    "7" {
        Write-Host "Tools: Azure AD Audit Sign In Log Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-AzureADAuditSignInLogSearch.ps1
        Start-AzureADAuditSignInLogSearch
    }

    "8" {
        Write-Host "Tools: Find all users with a specific RBAC Role" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-FindUserWithSpecificRbacRole.ps1
    }

    "9" {
        Write-Host "Tools: Find all users with all RBAC Role" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-AllUsersWithAllRoles.ps1
    }
    
    "10" {
        Write-Host "Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox" -ForegroundColor Green
        Start-Sleep -Seconds 3
        . $script:modulePath\ActionPlans\Start-MailboxDiagnosticLogs.ps1
    }
     
    "11" {
        Write-Host "Tools: Decode SafeLinks URL" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-DecodeSafeLinksURL.ps1
    }

    "12" {
        Write-Host "Tools: Export Quarantine Message" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Export-ExoQuarantineMessages.ps1
    }

    "13" {
        Write-Host "Tools: Transform IMCEAEX (old LegacyExchangeDN) to X500 address" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Get-X500FromImceaexNDR.ps1
    }

    "14" {
        Write-Host "DL to O365 Groups Upgrade Checker" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-DlToO365GroupUpdateChecks.ps1
    }

    "Q" {
        Write-Host "Quitting" -ForegroundColor Green
        Write-Host "All logs and files have been saved on $global:WSPath"
        Start-Sleep -Seconds 2
        Disconnect-all 
        #exit
        [Environment]::Exit(1)
    }
     
    default {
        Write-Host "I don't understand what you want to do. Will reload the menu!" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Clear-Host
        Start-O365TroubleshootersMenu 
     }
    } 
}