

#region Common Script Blocks
# Getting Credentials script block 
$Global:UserCredential = {
    Write-Host "`nPlease enter Office 365 Global Admin credentials:" -ForegroundColor Cyan
    $Global:O365Cred = Get-Credential
}

# Credential Validation block
$Global:CredentialValidation = { 
    If (!([string]::IsNullOrEmpty($errordescr)) -and !([string]::IsNullOrEmpty($global:error[0]))) {
        Write-Host "`nYou are NOT connected succesfully to $Global:banner. Please verify your credentials." -ForegroundColor Yellow
        $CurrentDescription = "`"" + $CurrentError + "`""
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

Function Connect-O365PS {
    # Function to connecto to O365 services

    # Parameter request and validation
    param (
        [ValidateSet("Msol", "AzureAd", "AzureAdPreview", "Exo", "ExoBasic", "Exo2", "Eop", "Scc", "AIPService", "Spo", "Sfb", "Teams", "ADSync")][Parameter(Mandatory = $true)] 
        $O365Service,
        [boolean] $requireCredentials = $True
    )
    $Try = 0
    $global:errordesc = $null
    $Global:O365Cred = $null
    . $script:modulePath\ActionPlans\Connect-ServicesWithTokens.ps1 

    
    #region Module Checks

    # Azure Module is mandatory

    if ((!($global:addTypeAzureAD) -and ("AzureAd" -in $O365Service ))) {
        $minimumVersionAzureAD = '2.0.2.16'
        $CurrentProperty = "Checking AzureAD Module"
        if ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD }).count -gt 0) {
            $pathModule = split-path ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD } | Sort-Object -Property version -Descending)[0]).Path -parent
            
            $path = join-path $pathModule 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
            try {
                Add-Type -Path $path
            }
            catch {}
            
            $CurrentDescription = "Azure AD Module for Windows PowerShell is installed and version is $(Split-Path $pathModule -Leaf)"
            $global:addTypeAzureAD = $true
        
        }
        else {
            $CurrentDescription = "Azure AD Module for Windows PowerShell is not installed or version is less than $minimumVersionAzureAD. Initiated install from PowerShell Gallery"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            #Uninstall-Module AzureAD -Force -Confirm:$false -ErrorAction SilentlyContinue |Out-Null
            
            Install-Module AzureAD -Force -Confirm:$false -AllowClobber
            $pathModule = split-path ((Get-Module azuread -ListAvailable | Where-Object { $_.Version -ge $minimumVersionAzureAD } | Sort-Object -Property version -Descending)[0]).Path -parent
            $path = join-path $pathModule 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll'
            try {
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

        If ((get-module -ListAvailable -Name MSOnline).count -eq 0 ) {
            $updateMSOL = $true
        }
        else {
            $updateMSOL = $true
            foreach ($version in (get-module -ListAvailable -Name MSOnline).Version) {
                if ($version -ge $minimumVersion) {
                    $updateMSOL = $false
                }
            }
        }
        if ($updateMSOL) {
            $CurrentProperty = "Checking MSOL Module"
            Write-Host "`nMSOL Module for Windows PowerShell is not installed. Initiated install from PowerShell Gallery" -ForegroundColor Red
            $CurrentDescription = "MSOL Module for Windows PowerShell is not installed or is less than required version $minimumVersion. Initiated install from PowerShell Gallery"
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Uninstall-Module MSOnline -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
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

        If ((get-module -ListAvailable -Name AzureADPreview).count -eq 0 ) {
            $updateAzureADPreview = $True
        }
        else {
            $updateAzureADPreview = $True
            foreach ($version in (get-module -ListAvailable -Name AzureADPreview).Version) {
                if ($version -ge $minimumVersion) {
                    $updateAzureADPreview = $false
                }
            }
        }
        if ($updateAzureADPreview) {
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

    if (($O365Service.tolower() -eq "exo2") -or ($O365Service.tolower() -eq "exo") -or ($O365Service.tolower() -eq "scc")) {
        if ((Get-Module -ListAvailable -Name ExchangeOnlineManagement).count -eq 0) {
            $CurrentProperty = "Checking ExchangeOnlineManagement v2 Module"
            $CurrentDescription = "ExchangeOnlineManagement module is not installed. We'll install it to support connectin to Exchange Online Module v2"
            write-host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Install-Module -Name ExchangeOnlineManagement -Force -Confirm:$false -AllowClobber
        }
    }

    # TODO: SPO prerequisites & modern module check

    # TODO: SFB prerequisites & modern module check

    If ( $O365Service -eq "ADSync") {
        If ((Get-Module -ListAvailable -Name ADSync).count -eq 0) {
            $CurrentProperty = "Checking ADSync Module"
            $CurrentDescription = "This dianostic have to be executed on AAD Connect server to have access to ADSync PowerShell Module"
            Write-Host "`n$CurrentDescription" -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Write-Host "The script was not executed from AAD Connect server." -ForegroundColor Red
            Write-Host "Returning to the main menu" -ForegroundColor Red
            Read-Key
            Start-O365TroubleshootersMenu
        }
    }

    if ($requireCredentials) {
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
                    $proxyConnection = "http://" + (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer

                }

                else {
                    #It doesn't work in PowerShell7
                    $proxyConnection = ([System.Net.WebProxy]::GetDefaultProxy()).Address.ToString()
                }
                Invoke-WebRequest -Proxy $proxyConnection  -ProxyUseDefaultCredentials https://provisioningapi.microsoftonline.com/provisioningwebservice.svc
            }
    
        }
    }
    #endregion Module Checks

    #region Connection scripts region
    if ($requireCredentials) {

        if ($global:MfaOption -eq 0) {
            do {
                # use only default modern authentication window (temporary disable manual managed tokens)
                #Write-Host "Do your account require MFA to authenticate? (y/n): " -ForegroundColor Cyan -NoNewline
                #$mfa = Read-Host 
                $mfa = "y"
                $mfa = $mfa.ToLower()
                if ($mfa -eq "y") {
                    $global:MfaOption = 1
                    Write-Host $global:MfaDisclaimer -ForegroundColor Red 
                    $global:userPrincipalName = Get-ValidEmailAddress("UserPrincipalName used to connect to Office 365 Services")
                }
                if ($mfa -eq "n") {
                    $global:MfaOption = 2
                    $global:credentials = Get-Credential -Message "Please input your Office 365 credentials:"
                    $global:userPrincipalName = $global:credentials.UserName
                }
            } until (($mfa -eq "y") -or ($mfa -eq "n"))
    

        }
    }
    switch ($O365Service) {
        # Connect to MSOL
        "MSOL" {
            # Defining the banner variable and clear the errors
            $Global:Error.Clear();
            $Global:banner = "MSOL PowerShell"
            $CurrentProperty = "Connect MSOL"
        
            if ($global:MfaOption -eq 2) {
                if (!(Get-Module MSOnline)) {
                    Import-Module MSOnline -Global -DisableNameChecking  -ErrorAction SilentlyContinue | Out-Null
                }
                $token = Get-TokenFromCache("AzureGraph")
                if ($null -eq $token) {
                    $token = Get-Token("AzureGraph")
                    Connect-MsolService -AdGraphAccessToken $token.AccessToken
                }
                else {
                    try {
                        $null = Get-MsolCompanyInformation -ErrorAction Stop
                    }
                    catch {
                        Connect-MsolService -AdGraphAccessToken $token.AccessToken
                    }
                }
            }
            elseif ($global:MfaOption -eq 1) {
                Do {

                    $errordescr = $null
                    $try++
                    try {
                        $null = Get-MsolCompanyInformation -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("MSOnline" -in (Get-Module).name)) {
                            Import-Module MSOnline -Global -DisableNameChecking  -ErrorAction SilentlyContinue | Out-Null
                        }
                        $errordescr = $null
                        Connect-MsolService -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        if ($null -eq $Global:Domain) {
                            $Global:Domain = (get-msoldomain -ErrorAction SilentlyContinue -ErrorVariable errordescr | Where-Object { $_.name -like "*.onmicrosoft.com" } | Where-Object { $_.name -notlike "*mail.onmicrosoft.com" }).Name
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
            if ($global:MfaOption -eq 2) {
                if (!(Get-Module AzureAD)) {
                    Import-Module AzureAD -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
                }
                $token = Get-TokenFromCache("AzureGraph")
                if ($null -eq $token) {
                    $token = Get-Token("AzureGraph")
                    Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue | Out-Null
                }
                else {
                    try {
                        $null = Get-AzureADTenantDetail -ErrorAction Stop
                    }
                    catch {
                        Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue | Out-Null
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
                    try {
                        $null = Get-AzureADTenantDetail -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("AzureAD" -in (Get-Module).name)) {
                            Import-Module AzureAD -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                        }
                        $errordescr = $null
                        Connect-AzureAd -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        if ($null -eq $Global:Domain) {
                            $Global:Domain = (Get-AzureADDomain -ErrorAction SilentlyContinue -ErrorVariable errordescr | Where-Object { $_.name -like "*.onmicrosoft.com" } | Where-Object { $_.name -notlike "*mail.onmicrosoft.com" }).Name
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
            if ($global:MfaOption -eq 3) {
                if (!(Get-Module AzureADPreview)) {
                    Import-Module AzureADPreview -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
                }
                $token = Get-TokenFromCache("AzureGraph")
                if ($null -eq $token) {
                    $token = Get-Token("AzureGraph")
                    AzureADPreview\Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue | Out-Null
                }
                else {
                    try {
                        $null = AzureADPreview\Get-AzureADTenantDetail -ErrorAction Stop
                    }
                    catch {
                        AzureADPreview\Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $global:userPrincipalName -ErrorVariable errordescr -ErrorAction SilentlyContinue | Out-Null
                    }
                }
            }
            elseif (($global:MfaOption -eq 1) -or ($global:MfaOption -eq 2)) {
                Do {
                    $errordescr = $null
                    $try++
                    try {
                        $null = AzureADPreview\Get-AzureADTenantDetail -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("AzureADPreview" -in (Get-Module).name)) {
                            Import-Module AzureADPreview -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
                        }
                        $errordescr = $null
                        AzureADPreview\Connect-AzureAD -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        if ($null -eq $Global:Domain) {
                            $Global:Domain = (AzureADPreview\Get-AzureADDomain -ErrorAction SilentlyContinue -ErrorVariable errordescr | Where-Object { $_.name -like "*.onmicrosoft.com" } | Where-Object { $_.name -notlike "*mail.onmicrosoft.com" }).Name
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

            if ($global:MfaOption -eq 2) {
                $token = Get-TokenFromCache("EXO")
                if ($null -eq $token) {
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
                else {
                    try {
                        $null = Get-OrganizationConfig -ErrorAction Stop
                    }
                    catch {
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
            elseif ($global:MfaOption -eq 1) {
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                Do {
                

                    $try++

                    try {
                        $null = Get-OrganizationConfig -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("ExchangeOnlineManagement" -in (Get-Module).Name)) {
                            Import-Module ExchangeOnlineManagement -Global -DisableNameChecking -Force -ErrorAction SilentlyContinue
                        }
                        
        
                        $errordescr = $null
                        if (($null -eq $Global:EXOSession ) -or ($Global:EXOSession.State -eq "Closed") -or ($Global:EXOSession.State -eq "Broken")) {
                            
                            Connect-ExchangeOnline -UserPrincipalName $global:UserPrincipalName -PSSessionOption $PSsettings -ShowBanner:$false -ErrorVariable errordescr -ErrorAction Stop 
                            $Global:EXOSession = Get-PSSession  | Where-Object { ($_.name -like "ExchangeOnlineInternalSession*") -and ($_.ConnectionUri -like "*outlook.office365.com*")  -and ($_.state -eq "Opened") }
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
                $Global:Domain = Get-AcceptedDomain | Where-Object { $_.name -like "*.onmicrosoft.com" } | Where-Object { $_.name -notlike "*mail.onmicrosoft.com" }  

            }
            catch {
                $CurrentDescription = "`"" + $CurrentError.ErrorRecord.Exception + "`""
            } 
            &$Global:DisplayConnect
        }
        # Connect to EXO2
        "EXO2" {

            $CurrentProperty = "Connect EXOv2"
            Do {
                # Defining the banner variable and clear the errors
                $Global:Error.Clear();
                $Global:banner = "EXOv2 PowerShell"
                $errordescr = $null
                $try++
                try {
                    $null = Get-EXOMailbox -ErrorAction Stop
                }
                catch {
                    Write-Host "$CurrentProperty"
                    if (!("ExchangeOnlineManagement" -in (Get-Module).name)) {
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

            if ($global:MfaOption -eq 2) {
                $token = Get-TokenFromCache("EXO")
                if ($null -eq $token) {
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
                else {
                    try {
                        $null = Get-OrganizationConfig -ErrorAction Stop
                    }
                    catch {
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
            elseif ($global:MfaOption -eq 1) {
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
                    $Global:Domain = Get-AcceptedDomain | Where-Object { $_.name -like "*.onmicrosoft.com" } | Where-Object { $_.name -notlike "*mail.onmicrosoft.com" } 
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

            if ($global:MfaOption -eq 2) {
                $token = Get-TokenFromCache("EXO")
                if ($null -eq $token) {
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
                else {
                    try {
                        $null = Get-OrganizationConfig -ErrorAction Stop
                    }
                    catch {
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
            elseif ($global:MfaOption -eq 1) {
                Do {
                    # Defining the banner variable and clear the errors
                    $Global:Error.Clear();
                    $Global:banner = "Security&Compliance Online PowerShell - Modern & MFA"
                    $try++

                    try {
                        $null = Get-OrganizationConfig -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("ExchangeOnlineManagement" -in (Get-Module).Name)) {
                            Import-Module ExchangeOnlineManagement -Global -DisableNameChecking -Force -ErrorAction SilentlyContinue
                        }
                        
        
                        $errordescr = $null
                        if (($null -eq $Global:EXOSession ) -or ($Global:EXOSession.State -eq "Closed") -or ($Global:EXOSession.State -eq "Broken")) {
                            
                            Connect-IPPSSession -UserPrincipalName $global:UserPrincipalName -PSSessionOption $PSsettings -ShowBanner:$false -ErrorVariable errordescr -ErrorAction Stop 
                            $Global:EXOSession = Get-PSSession  | Where-Object { ($_.name -like "ExchangeOnlineInternalSession*") -and ($_.ConnectionUri -like "*compliance.protection.outlook.com*") -and ($_.state -eq "Opened") }
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
                        $URL = read-host "Please Input the connection URL (i.e.: https://Tenant_Domain-admin.sharepoint.com/)"
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
        
            if ($global:MfaOption -eq 2) {
                if (!(Get-Module AIPService)) {
                    Import-Module AIPService -Global -DisableNameChecking  -ErrorAction SilentlyContinue 
                }
                $token = Get-TokenFromCache("AIPService")
                if ($null -eq $token) {
                    $token = Get-Token("AIPService")
                    Connect-AipService -AccessToken $token.AccessToken -ErrorAction SilentlyContinue | Out-Null
                }
                else {
                    try {
                        $null = Get-AipServiceConfiguration -ErrorAction Stop
                    }
                    catch {
                        Connect-AipService -AccessToken $token.AccessToken -ErrorAction SilentlyContinue | Out-Null
                    }
                }
            }
            elseif ($global:MfaOption -eq 1) {
                do {
                    $Global:Error.Clear();
                    $Global:banner = "AIPService PowerShell"
                    $errordescr = $null
                    $try++
                    try {
                        $null = Get-AipServiceConfiguration -ErrorAction Stop
                    }
                    catch {
                        Write-Host "$CurrentProperty"
                        if (!("AIPService" -in (Get-Module).name)) {
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
        "ADSync" {

            try {

                Import-Module  ADSync -DisableNameChecking -Global -ErrorAction SilentlyContinue
                $CurrentDescription = "Success"

            }
            catch {
                $CurrentDescription = "`"" + $CurrentError.ErrorRecord.Exception + "`""
            } 
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
    $global:Path = [Environment]::GetFolderPath("Desktop")
    $Global:Path += "\PowerShellOutputs"
    $global:WSPath = "$Path\PowerShellOutputs_$ts"
    $global:starline = New-Object String '*', 5
    $global:addTypeAzureAD = $false
    $global:MfaOption = 0;
    #$Global:ExtractXML_XML = "Get-MigrationUserStatistics ", "Get-ImapSubscription "
    $global:Disclaimer = 'Note: Before you run the script: 

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
    
    if (!(Test-Path $Path)) {
        Write-Host "We are creating the following folder $Path"
        New-Item -Path $Path -ItemType Directory -Confirm:$False | Out-Null
    }

    if (!(Test-Path $WSPath)) {
        Write-Host "We are creating the following folder $WSPath"
        New-Item -Path $WSPath -ItemType Directory -Confirm:$False | Out-Null
    }
    
    $global:outputFile = "$WSPath\Log_$ts.csv"
    $global:columnLabels = "Time, Function, Step, Description"
    Out-File -FilePath $outputFile -InputObject $columnLabels -Encoding UTF8 | Out-Null

    Set-Location $WSPath
    Write-Host "`n"

    #if ($null -eq $global:credential)
    #{
    #$global:userPrincipalName = Get-ValidEmailAddress("UserPrincipalName used to connect to Office 365 Services")
    #Write-Host "Please note that depening the Office 365 Services we need to connect, you might be asked to re-add the UserPrincipalName in another Authentication Form!" -ForegroundColor Yellow
    #Start-Sleep -Seconds 5
    #}

    
}

function Get-ValidEmailAddress([string]$EmailAddressType) {
    [int]$count = 0
    do {
        Write-Host "Enter Valid $EmailAddressType`: " -ForegroundColor Cyan -NoNewline
        [string]$EmailAddress = Read-Host
        [bool]$valid = ($EmailAddress.Trim() -match "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$")
        if (!$valid) {
            $InvalidEmailAddressWarning = "Inputed Email Address `"$EmailAddress`" does not pass O365Troubleshooters format validation"
            Write-Warning -Message $InvalidEmailAddressWarning
            Write-Log -function "Get-ValidEmailAddress" -step "input address" -Description $InvalidEmailAddressWarning
        }
        $count++
    }
    while (!$valid -and ($count -le 2))
    
    if ($valid) {
        return $EmailAddress.Trim()
    }
    else {   
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
    param ([string][Parameter(Mandatory = $true)]$IntType)

    [int]$count = 0

    do {
        [bool]$valid = $true
        Write-Host "Enter Valid $IntType`: " -ForegroundColor Cyan -NoNewline
        try {
            [int]$IntFromConsole = Read-Host
            if ($IntFromConsole -eq 0) {
                throw "System.Management.Automation.RuntimeException"
            }
        }
        catch [System.Management.Automation.RuntimeException] {
            Write-Host "Invalid $IntType returned" -ForegroundColor Red
            $valid = $false
            $count++
        }
    }while (!$valid -and ($count -le 2))
    
    if ($valid) {
        return $IntFromConsole
    }
    else {   
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
              
            $CurrentDescription = "`"" + $CurrentDescription + "`"" 
                
        }
        write-log -Function "New-XMLObject" -Step $CurrentProperty -Description $CurrentDescription
        
        $myerror = $null
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
    #Write-Host

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
    
    [int]$i = 0
    do {
        
        $OptionsList | ForEach-Object {
            write-host "Press $($_[0]) for '$_'"
        }
        [string]$Option = read-host "Please answer by typing first letter of the option" 
        [bool]$validAnswer = $false
        $OptionsList | ForEach-Object {
            if ($_.ToLower()[0] -eq $Option.ToLower()[0]) {
                $validAnswer = $true
            }
        }
        $i++
    }
    while (($validAnswer -eq $false) -and ($i -le 2))
    if ($validAnswer -eq $false) {
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

    if (Test-Path $v4Directory) {
        $version = Get-ItemProperty $v4Directory -name Version | Select-Object -expand Version
        $dotnet = ($version).Split(".")
        If (($dotnet[0] -eq 4) -and ($dotnet[1] -ge 5)) {
            Write-Host "You have the following .NET Framework version installed: " $version
            Write-Host "The .NET Framework version meets the minimum requirements" -foregroundcolor "green"
        }
        else {
            Write-Host "You have the following .NET Framework version installed: " $version
            Write-Host "Your .net version is less than 4.5. Please update the .NET Framework version" -foregroundcolor "red"
            Open-URL ("http://go.microsoft.com/fwlink/?LinkId=671744")
            Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
            exit
        }
    }
    else {
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
    If ($PSVersionTable.PSVersion.Major -le 2) {
        If ($op_ver_srv) {
            Write-Host "`nYou have a Server Operating System !" -foregroundcolor "red"
            write-log -Function "Test-PSVers" -step "Check Powershell version" -Description $PSVersionTable.PSVersion.Major
            Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
            exit
        }
        else {
            Test-DotNet
            Write-Host "`nYou have the following Powershell version:" $PSVersionTable.PSVersion.Major
            Write-Host "`nYour Powershell version is less than 5. Please update your Powershell by installing Windows Management Framework 5.0 !" -foregroundcolor "magenta"
            Open-URL ("https://www.microsoft.com/en-us/download/details.aspx?id=50395")
            Write-Host "`nThe Collection script will now stop" -foregroundcolor "red"
            exit
        }
    }
    Else {
        Write-Host "`nYou have the following Powershell version:" $PSVersionTable.PSVersion.Major -foregroundcolor "green"
    }

}

function Start-Elevated {
    
    If (!([Net.ServicePointManager]::SecurityProtocol -eq [Net.SecurityProtocolType]::Tls12 ))
    {
        #Bypass the question as anyway the configuration is just per session, won't persist after restart
        #write-host "SecurityProtocol version should be TLS12 for PowerShellGet to be installed. If the value will different than TLS12, the script will exit" -ForegroundColor Red
        #$answer = Read-Host "Do you agree to set SecurityProtocol to Tls12? Type y for `"Yes`" and n for `"No`""
        $answer ="y"
        if ($answer.ToLower() -eq "y")
        {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
            Write-Host "SecurityProtocol has been set to TLS12!" -ForegroundColor Green
        }
        else
        {
            Write-Host "As you did't choose to set the value to TLS12, the script will exit!" -ForegroundColor Red
            Read-Key
            Exit
        }
    }

    Write-Host "Starting new PowerShell Window with the O365Troubleshooters Module loaded"
    Read-Key
    Start-Process powershell.exe -ArgumentList "-noexit -Command Install-Module O365Troubleshooters -force; Import-Module O365Troubleshooters -force; Start-O365Troubleshooters -elevatedExecution `$true" -Verb RunAs #-Wait
    #Exit
}

function Disconnect-All {
    
    $CurrentDescription = "Disconnect is successful!"

    try {
        # Disconnect EXOv2
        if (("ExchangeOnlineManagement" -in (Get-Module).name)) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Disconnect AIP
        if (("AipService" -in (Get-Module).name)) {
            Disconnect-AipService -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Disconnect AzureAD
        if (("AzureAD" -in (Get-Module).name)) {
            AzureAD\Disconnect-AzureAD -Confirm:$false -ErrorAction SilentlyContinue
            Remove-Module AzureAD -Force -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Disconnect AzureADPreview
        if (("AzureADPreview" -in (Get-Module).name)) {
            AzureADPreview\Disconnect-AzureAD -Confirm:$false -ErrorAction SilentlyContinue
            Remove-Module AzureADPreview -Force -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Disconnect all PsSessions
        Get-PSSession | Remove-PSSession
    }

    catch {
             
        $CurrentDescription = "`"" + $Global:Error[0].Exception.Message + "`"" 

    }

    write-log -Function "Disconnect - close sessions" -Step $CurrentProperty -Description $CurrentDescription

    
    $CurrentDescription = "Execution Policy was successfully set to its original value!"

    try {
        # Set back the initial ExecutionPolicy value
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy $_CurrentUser -Force -ErrorAction SilentlyContinue
    }
    catch {
        $CurrentDescription = "`"" + $Global:Error[0].Exception.Message + "`"" 
    }
    
    write-log -Function "Disconnect - ExecutionPolicy" -Step $CurrentProperty -Description $CurrentDescription
    # Read-Host -Prompt "Please press [Enter] to continue"
    Read-Key
}

Function Read-Key {
    if ($psISE) {
        Write-Host "Press [Enter] to continue" -ForegroundColor Cyan
        Read-Host 
    }
    else {
        Write-Host "Press any key to continue" -ForegroundColor Cyan
        #[void][System.Console]::ReadKey($true)
        $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
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
###             [string]$DataType - Contains the data type of the data that need to be added into the HTML file (it can be: "CustomObject" (aka a "Table"), "String")
###             [CustomObject]/[String]$EffectiveData - Contains the effective data that need to be added into a HTML file
###             [String]$TableType - If EffectiveData is table, we need to know how to list it (accepted values: "List" = Vertical, "Table" = Horizontal)
### </param>
function Export-ReportToHTML {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [string]$PageTitle,

        [Parameter(Mandatory = $false)]
        [string]$ReportTitle,

        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$TheObjectToConvertToHTML
    )

    ### Create header of the HTML file
    $HTMLBeginning = @"
<!DOCTYPE html>
<html>
	<head>
	
		<meta charset="UTF-8">
		<meta name="description" content="Office 365 Troubleshooters">
		<meta name="keywords" content="Office 365, Troubleshooters">
		<meta name="author" content="Cristian Dimofte">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		
        <title>$PageTitle</title>
		<link rel="titlebar icon" href="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABd0lEQVQ4je2RTyikcRjHP887LzPj/yRli1DL+puDcpLS0rZ7cqOUcnDblqubg3JQcsARSU5KUVKcRLurzcWMmcVhylXNGMQYv/dxeLGD5Kz2e3t6vs+n7/M88F/ylqFhQUORGOq3kd5qrMk2qXkTsHysB7YH6V7DXKae9vrqkNkOqXsVkD2lwedD6fLakPwu9Q+1ld6snddgvjfN7MHpqeaSQakfb+U800Ir8pCOZQ0N/9TwkwRD2xoO+HEaC7G61jGJJOoMSMO3Ff1dnIU90y5Nnau6/7UM++QCHdnFqSpALIDcaQ2O/sFcpZAv5VKz2YmnyOeCN6JkJZJu0kgM/ZCDZGcgJbnIYRy1Kuc0dH4DAR8U+bEm9jQcTcC1cQGatuJJAoZ2MHmZMNbiQu22EuQojn4uRX5scXtrHv3ycVZ/dX8i1lyMdwlwgKM4enqFzEcwXvv+Bv2berD4F+fG8EIiEPCROrvGNur6LYHSHOS4798737HuADOWgsu9EJJHAAAAAElFTkSuQmCC" />
		
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />
		<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>
		
		<style type="text/css">
			html {
				box-sizing: border-box;
				font-family: FabricMDL2Icons;
			}
			
			*, *:before, *:after {
				box-sizing: inherit;
			}			
			
			.ms-font-su {
				font-family:FabricMDL2Icons;
				-webkit-font-smoothing:antialiased;
				font-weight:100
			}
			
			.ms-fontColor-themePrimary,.ms-fontColor-themePrimary--hover:hover{
				color:#0078d7
			}
			
			@font-face{
				font-family:FabricMDL2Icons;
				src:url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.woff) format('woff'),url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.ttf) format('truetype');
				font-weight:400;
				font-style:normal
			}

			body {
				padding: 10px;
				font-family: FabricMDL2Icons;
				background: #f6f6f6;
			}
			
			a {
				color: #06c;
			}

			h1 {
				margin: 0 0 1.5em;
				font-weight: 600;
				font-size: 1.5em;
			}

			h2 {
				color: #1a0e0e;
				font-size: 20px;
				font-weight: 700;
				margin: 0;
				line-height: normal;
				text-transform:uppercase;
				margin-block-start: 0.3em;
				margin-block-end: 0.3em;
				text-indent: 0px;
				margin-left: 0px;
			}
    
			h2 span {
				display: block;
				padding: 0;
				font-size: 18px;
				opacity: 0.7;
				margin-top: 5px;
				text-transform:uppercase;
			}

			h3 {
				color: #000000;
				font-size: 17px;
				font-weight: 300;
				margin-left: 10px;
				font-family: FabricMDL2Icons;
			}
    
			h3 span {
				color: #000000;
				font-size: 17px;
				font-weight: 300;
				margin-left: 10px;
				font-family: FabricMDL2Icons;
			}			
			
			h5 {
				display: block;
				padding: 0;
				margin: 5px 0px 0px 0px;
				font-size: 42px;
				font-weight: 100;
				font-family: FabricMDL2Icons;
			}
			
			p {
				margin: 0 0 1em;
				padding: 10px;
				font-family: FabricMDL2Icons;
			}

			.label {
				width: 100%;
				height: 80px;
				background-color: #0078d7;
				color: #ffffff;
				font-size: 46px;
				display: inline-block;
			}
			
			.label1 {
				width: 100%;
				height: 80px;
				background-color: #f6f6f6;
				color: #0078d7;
				font-size: 46px;
				display: inline-block;
				box-sizing: border-box;
			}
			
			.body-panel {
				font-size: 14px;
				padding: 15px;
				margin-bottom: 5px;
				margin-top: 5px;
				box-shadow: 0 1px 2px 0 rgba(0,0,0,.1);
				overflow-x: auto;
				box-sizing: border-box;
				background-color: #fff;
				color: #333;
				-webkit-tap-highlight-color: rgba(0,0,0,0);
				padding-top: 15px;
				font-family: FabricMDL2Icons;
			}

			.accordion {
				margin-bottom: 1em;
			}

			.accordion p:last-child {
				margin-bottom: 0;
			}
			
			.accordion > input[name="collapse"] {
				display: none;
			}
			
			.accordion > input[name="collapse"]:checked ~ .content {
				height: auto;
				transition: height 0.5s;
			}
			
			.accordion label, .accordion .content {
				max-width: 3200px;
				width: 99%;
			}
			
			.accordion .content {
				background: #fff;
				overflow: hidden;
				overflow-x: auto;
				height: 0;
				transition: 0.5s;
				box-shadow: 1px 2px 4px rgba(0, 0, 0, 0.3);
			}
			
			.accordion label {
				display: block;
			}
			
			.accordion > input[name="collapse"]:checked ~ .content {
				border-top: 0;
				transition: 0.3s;
			}
			
			.accordion .handle {
				margin: 0;
				font-size: 16px;
			}
			
			.accordion label {
				color: #0078d7;
				cursor: pointer;
				font-weight: 300;
				padding: 10px;
				background: #e6f3ff;
				user-select: none;
			}
			
			.accordion label:hover, .accordion label:focus {
				background: #cce6ff;
				color: #0000ff;
			}
			
			.accordion .handle label:before {
				font-family: FabricMDL2Icons;
				content: "\E972";
				display: inline-block;
				margin-right: -20px;
				font-size: 1em;
				line-height: 1.556em;
				vertical-align: middle;
				transition: 0.4s;
			}
			
			.accordion > input[name="collapse"]:checked ~ .handle label:before {
				transform: rotate(180deg);
				transform-origin: center;
				transition: 0.4s;
			}
			
			section {
				float: left;
				width: 100%;
			}
    
			.container{
				max-width: 3200px;
				width:99%;
				margin: 0 auto;
			}

		   table {
				font-size: 14px;
				font-weight: 800;
				border: 1px solid #0078d7;
				border-collapse: collapse;
				font-family: FabricMDL2Icons;
				margin-left: 10px;
				margin-bottom: 10px;
				text-align: center;
			} 
			
			td {
				padding: 4px;
				margin: 0px;
				border: 1px solid #0078d7;
				border-collapse: collapse;
			}
			
			th {
				color: #0078d7;
				font-family: FabricMDL2Icons;
				text-transform: uppercase;
				padding: 10px 5px;
				vertical-align: middle;
				border: 1px solid #0078d7;
				border-collapse: collapse;
			}
			
			tbody tr:nth-child(odd) {
				background: #f0f0f2;
			}			

			#output {
				color: #004377;
				font-size: 14px;
			}
			
			@media screen and (max-width:639px){
				h2 {
					color: #1a0e0e;
					font-size: 16px;
					font-weight: 700;
					margin: 0;
					line-height: normal;
					text-transform:uppercase;
					text-indent: 0px;
					margin-left: 0px;
				}
			
				h5 {
					margin: 5px 0px 0px 0px;
					font-size: 8vw;
					font-weight: 100;
					font-family: FabricMDL2Icons;
				}
				
				.accordion label, .accordion .content {
					max-width: 3200px;
					width: 99%;
				}
				
				.accordion .content {
					background: #fff;
					overflow: hidden;
					overflow-x: auto;
					height: 0;
					transition: 0.5s;
					box-shadow: 1px 2px 4px rgba(0, 0, 0, 0.3);
				}
				
				.accordion > input[name="collapse"]:checked ~ .content {
					height: auto;
				}
			}
			
			@media (max-device-width:480px) and (orientation:landscape) {
				.body-panel {
					max-height:200px
				}
			}

			/* For Desktop */
			@media only screen and (min-width: 620px) {
				.accordion > input[name="collapse"]:checked ~ .content {
					height: auto;
				}
			}	
		</style>
		
	</head>
	<body class="ms-Fabric" dir="ltr">
		<header>
		<div class="label1">
			<a href="https://www.powershellgallery.com/packages/O365Troubleshooters" target="_blank">
				<img src="data:image/png;base64,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" alt="O365Troubleshooters" width="auto" height="87%" style="float: left; margin: 5px 10px">
			</a>
			<div style="width: 100%; height: 100%; ">
			
                <h5 class="ms-font-su ms-fontColor-themePrimary" style="font-weight: 350; ">
                    $ReportTitle
                </h5>
            </div>
        </div>
        </header>
        <div class="body-panel" style="margin: 30px 0px 0px 0px; display: block;">
"@

    $HTMLEnd = @"
		</div>
		
		<div class="ms-font-su ms-fontColor-themePrimary" style="font-size: 15px; margin-left: 10px; text-align: right;">
			<ul>Creation Date: $((Get-date).ToUniversalTime()) UTC</ul>
			<ul>&copy; 2021 O365Troubleshooters</ul>
		</div>
	</body>
</html>
"@


    [int]$i = 1

    [string]$TheBody = $HTMLBeginning

    ### For each scenario, convert the data to HTML
    foreach ($Entry in $TheObjectToConvertToHTML) {
        $TheBody = $TheBody + "
        	`<section class=`"accordion`"`>
				`<input type=`"checkbox`" name=`"collapse`" id=`"handle$i`" `>
				`<h2 class=`"handle`"`>
					`<label for=`"handle$i`" `>
						`<h2 class=`"ms-font-su ms-fontColor-themePrimary`" style=`"display: inline-block; color: $($Entry.SectionTitleColor); font-size: 20px; font-weight: 650;`"`>&nbsp;&nbsp;&nbsp;&nbsp;$($Entry.SectionTitle)`</h2`>
					`</label`>
				`</h2`>
				`<div class=`"content`"`>
					`<h3`>$($Entry.Description)`</h3`>
        "

        if ($Entry.DataType -eq "String") {
            $TheValue = "					`<p style=`"font-family: FabricMDL2Icons; font-weight: 800; margin-left: 10px;`"`>$($Entry.EffectiveData)`<`/p>"
        }
        else {
            $TheProperties = ($($Entry.EffectiveData) | Get-Member -MemberType NoteProperty).Name
            $TheValue = $($Entry.EffectiveData) | ConvertTo-Html -As $($Entry.TableType) -Property $TheProperties -Fragment | ForEach-Object { (($_.Replace("&lt;", "<")).Replace("&gt;", ">")).replace("&quot;", '"') }
            
            if ($Entry.TableType -eq "List") {
                [int]$z = 0
                foreach ($NewEntryFound in $TheValue) {
                    if ($NewEntryFound -like "*<table>*") {
                        $TheValue.Item($z) = $NewEntryFound.Replace("<table>", "<table style=`"text-align: left;`">")
                    }
                    elseif ($NewEntryFound -like "*<tr><td>*") {
                        $TheValue.Item($z) = $NewEntryFound.Replace("<tr><td>", "<tr><th>")
                        $TheValue.Item($z) = (($TheValue.Item($z) -split "<td")[0] -replace "`/td>", "`/th>") + ($TheValue.Item($z).Substring((($TheValue.Item($z) -split "<td")[0].Length), ($TheValue.Item($z).Length - ($TheValue.Item($z) -split "<td")[0].Length)))
                    }
                    $z++
                }
            }

        }



        ### Adding sections in the body of the HTML report
        $TheBody = $TheBody + $TheValue
        $TheBody = $TheBody + "
        	`<`/div`>
		`<`/section`>
        "

        $i++
    }

    $TheBody = $TheBody + $HTMLEnd

    $TheBody | Out-File $FilePath -Force
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
        [Parameter(ParameterSetName = "String", Mandatory = $false)]
        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [string]$SectionTitle,

        [Parameter(ParameterSetName = "String", Mandatory = $false)]
        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [ValidateSet("Black", "Green", "Red")]
        [string]$SectionTitleColor,

        [Parameter(ParameterSetName = "String", Mandatory = $false)]
        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [string]$Description,

        [Parameter(ParameterSetName = "String", Mandatory = $false)]
        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [ValidateSet("CustomObject", "String")]
        [string]$DataType,

        [Parameter(ParameterSetName = "String", Mandatory = $false)]
        [string]$EffectiveDataString,

        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [PSCustomObject]$EffectiveDataArrayList,

        [Parameter(ParameterSetName = "CustomObject", Mandatory = $false)]
        [ValidateSet("List", "Table")]
        [string]$TableType
    )

    ###Create the object, with all needed Properties, that will be used to convert into an HTML report
    $TheObject = New-Object PSObject
    $TheObject | Add-Member -NotePropertyName SectionTitle -NotePropertyValue $SectionTitle
    $TheObject | Add-Member -NotePropertyName SectionTitleColor -NotePropertyValue $SectionTitleColor
    $TheObject | Add-Member -NotePropertyName Description -NotePropertyValue $Description
    $TheObject | Add-Member -NotePropertyName DataType -NotePropertyValue $DataType
    if ($DataType -eq "CustomObject") {
        $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataArrayList
        $TheObject | Add-Member -NotePropertyName TableType -NotePropertyValue $TableType
    }
    else {
        $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataString
    }

    ### Return the created object
    return $TheObject

}


Function    Start-O365Troubleshooters {
    param(
        [bool][Parameter(Mandatory = $false)] $elevatedExecution = $false
    )
    if (!$elevatedExecution) {
        Start-Elevated
    }
    else {
        If (!([Net.ServicePointManager]::SecurityProtocol -eq [Net.SecurityProtocolType]::Tls12))
        {
            write-host "SecurityProtocol version should be TLS12 for PowerShellGet to be installed. If the value will different than TLS12, the script will exit" -ForegroundColor Red
            $answer = Read-Host "Do you agree to set SecurityProtocol to Tls12? Type y for `"Yes`" and n for `"No`""
            if ($answer.ToLower() -eq "y")
            {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
                Write-Host "SecurityProtocol has been set to TLS12!" -ForegroundColor Green
            }
            else
            {
                Write-Host "As you did't choose to set the value to TLS12, the script will exit!" -ForegroundColor Red
                Read-Key
                Exit
            }
        }
        try {
            Set-GlobalVariables
            Start-O365TroubleshootersMenu
        }
        catch {
            # As we don't know at which step the unhundled exception occures, we test for output path first
            $errorGenericPath = "$([Environment]::GetFolderPath("Desktop"))\PowerShellOutputs\"
            if (!(Test-Path $errorGenericPath ))
            {
                Write-Host "We are creating the following folder $errorGenericPath "
                New-Item -Path $errorGenericPath -ItemType Directory -Confirm:$False | Out-Null
            }

            # Export the unhundled exception and log this event in log file, too
            Write-Host "Script Encountered an un-handled exception! This will exported in Folder: $([Environment]::GetFolderPath("Desktop"))\PowerShellOutputs\"
            $_ | Export-Clixml -Depth 100 -Path "$([Environment]::GetFolderPath("Desktop"))\PowerShellOutputs\PowerShellOutputs_UnhandledError_$(Get-Date -Format yyyyMMdd_HHmmss).xml"
            #TODO: should implement additional test here for Write-Log folder path
            Write-Log -function "ERROR" -step "Unhandled" -Description "ERROR: $($PSItem.Exception.Message)"
            Start-Sleep -Seconds 5
            Exit
        }

    }
}

Function Start-O365TroubleshootersMenu {
    $menu = @"
    1  Encryption: Office Message Encryption General Troubleshooting
    2  Mail Flow: SMTP Relay Test
    3  Migration: Analyze Mailbox move (Hybrid migration)
    4  Security: Compromised Tenant Investigation
    5  Groups: DL to O365 Groups Upgrade Checker
    6  DDG to Exchange Online Contact automatically synchronization with AAD Connect
    7  Public Folder Troubleshooter
    8  Tools: Exchange Online Audit Search
    9  Tools: Unified Logging Audit Search
    10 Tools: Azure AD Audit Sign In Log Search
    11 Tools: Find all users with a specific RBAC Role
    12 Tools: Find all users with all RBAC Roles
    13 Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox
    14 Tools: Decode SafeLinks URL
    15 Tools: Export Quarantine Messages
    16 Tools: Transform IMCEAEX (old LegacyExchangeDN) to X500 address
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
            Start-Office365Relay
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
            Write-Host "Action Plan: DL to O365 Groups Upgrade Checker" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-DlToO365GroupUpgradeChecks.ps1
        }

        "6" {
            Write-Host "Action Plan: DDG to Exchange Online Contact automatically synchronization with AAD Connect" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-SyncDDGasContactwithAADConnect.ps1
        }
        "7" {
            Write-Host "Public Folder Troubleshooter" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-PublicFolderTroubleshooter.ps1
        }
        "8" {
            Write-Host "Tools: Exchange Online Audit Search" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1
            Start-ExchangeOnlineAuditSearch
        }

        "9" {
            Write-Host "Tools: Unified Logging Audit Search" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-UnifiedAuditLogSearch.ps1
        }

        "10" {
            Write-Host "Tools: Azure AD Audit Sign In Log Search" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-AzureADAuditSignInLogSearch.ps1
            Start-AzureADAuditSignInLogSearch
        }

        "11" {
            Write-Host "Tools: Find all users with a specific RBAC Role" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-FindUserWithSpecificRbacRole.ps1
        }

        "12" {
            Write-Host "Tools: Find all users with all RBAC Role" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-AllUsersWithAllRoles.ps1
        }
    
        "13" {
            Write-Host "Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox" -ForegroundColor Green
            Start-Sleep -Seconds 3
            . $script:modulePath\ActionPlans\Start-MailboxDiagnosticLogs.ps1
        }
     
        "14" {
            Write-Host "Tools: Decode SafeLinks URL" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Start-DecodeSafeLinksURL.ps1
        }

        "15" {
            Write-Host "Tools: Export Quarantine Message" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Export-ExoQuarantineMessages.ps1
        }

        "16" {
            Write-Host "Tools: Transform IMCEAEX (old LegacyExchangeDN) to X500 address" -ForegroundColor Green
            . $script:modulePath\ActionPlans\Get-X500FromImceaexNDR.ps1
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