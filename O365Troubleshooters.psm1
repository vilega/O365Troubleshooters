

#region Common Script Blocks
    # Getting Credentials script block 
    $Global:UserCredential = {
            Write-Host "`nPlease enter Office 365 Global Admin credentials:" -ForegroundColor Cyan
            $Global:O365Cred = Get-Credential
    }

    # Credential Validation block
    $Global:CredentialValidation = { 
            If ($Global:Error -ne $null) {
                Write-Host "`nYou are NOT connected succesfully to $Global:banner. Please verify your credentials." -ForegroundColor Yellow
                $CurrentDescription = "`""+$CurrentError+"`""
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                &$Global:UserCredential
            }
    }

    # Displaying connection status
    $Global:DisplayConnect = {
            If ($Global:Error -ne $null) {
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
    param ([ValidateSet("msol","exo","exo2","eop","sco","spo","sfb","AIPService")][Parameter(Mandatory=$true)] 
            $O365Service 
    )
    $Try = 0
    $global:errordesc = $null
    $Global:O365Cred=$null
    
#region Module Checks
    # $O365Service = "MSOL", "EXO" - Checking if the Azure Active Directory Module for Windows PowerShell (64-bit version) modules are installed on the machine
    If ( $O365Service -match "MSOL") {
            If ((get-module -ListAvailable -Name *MSOnline*).count -eq 0 ) {
                $CurrentProperty = "CheckingMSOL Module"
                Write-Host "`nAzure Active Directory Module for Windows PowerShell is not installed. Please go to 'https://technet.microsoft.com/en-us/library/jj151815.aspx' in order to install is and and after that re-run the script" -ForegroundColor Yellow
                $CurrentDescription = "Azure Active Directory Module for Windows PowerShell is not installed. Please go to 'https://technet.microsoft.com/en-us/library/jj151815.aspx' in order to install is and and after that re-run the script"
                Write-Host "`nNow the script will stop." -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                Read-Host
                Exit
            }
    }

    # Checking if the Sharepoint Online PowerShell Module is installed on the machine
    If ( $O365Service -match "SPO") {
            If ((Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell).count -eq 0) {
                $CurrentProperty = "CheckingSPO Module"
                Write-Host "`nSharePoint Online Management Shell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=35588' in order to download and install the module" -ForegroundColor Yellow
                $CurrentDescription = "SharePoint Online Management Shell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=35588' in order to download and install the module. !!! Please restart your computer after the installtion finishes !!!"
                Write-Host "`nNow the script will stop." -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                Read-Host
                Exit
            } 
    }

    # Checking if the Skype for Business Online, Windows PowerShell Module is installed on the machine
    If ( $O365Service -match "SFB") {
            If ((Get-Module -ListAvailable -Name LyncOnlineConnector).count -eq 0) {
                $CurrentProperty = "CheckingSFB Module"
                Write-Host "`nSkype for Business Online, Windows PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=39366' in order to download and install the module" -ForegroundColor Yellow
                $CurrentDescription = "Skype for Business Online, Windows PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=39366' in order to download and install the module"
                Write-Host "`nNow the script will stop." -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                Read-Host
                Exit
            } 
    }

    If ( $O365Service -match "AIPService") {
        If ((Get-Module -ListAvailable -Name AIPService).count -eq 0) {
            $CurrentProperty = "Checking AIPService Module"
            
            Write-Host "AIPService needs to be updated or you have just updated without restarting the PC/laptop" -ForegroundColor Red
            Write-Host "We will try to install the AIPService module" -ForegroundColor Cyan
            Install-Module -Name AIPService -Force -Confirm:$false
            Write-Host "Installed the AADRM module"
            #Import-Module AIPService -Force

            #TODO: check if AADRM Module was succesfully installed
            <# 
            $CurrentDescription = "Azure Active Directory Right Management PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=30339' in order to download and install the module"
            Write-Host "`nNow the script will stop." -ForegroundColor Red
            write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
            Read-Host
            Exit
            #>
            #TODO: if AADRM module was installed, that needs to be uninstalled

        }
    }
    
    if ( $O365Service -match "exo2") 
    {
        If ((Get-Module -ListAvailable -Name ExchangeOnlineManagement).count -eq 0) 
        {
            $CurrentProperty = "Checking AIPService Module"
            write-host "ExchangeOnlineManagement module is not installed. We'll install it to support connectin to Exchange Online Module v2" -ForegroundColor Red
            Install-Module -Name ExchangeOnlineManagement -Force -Confirm:$false
        }
    }

   
    #$Global:proxy = Read-Host
    Write-Host "`nAre you able to access Internet from this location without a Proxy?" -ForegroundColor Cyan
    $Global:proxy = get-choice "Yes", "No"
    
    $Global:PSsettings = New-PSSessionOption -SkipRevocationCheck 
    if ($Global:proxy -eq "n") {
                $Global:PSsettings = New-PSSessionOption -ProxyAccessType IEConfig -SkipRevocationCheck 
    }

#endregion Module Checks

#region Connection scripts region
  switch ($O365Service) {
    # Connect to AzureAD (MSOL) PowerShell
    "MSOL" {
                $Global:Error.Clear();
                Import-Module MSOnline ;
                If ($null -eq $Global:O365Cred) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting 
                
                $CurrentProperty = "Connect MSOL"
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Azure AD (MSOL) PowerShell"
                        $try++
                        # Creating the session for PS MSOL Service
                        Connect-MsolService -Credential $O365Cred -ErrorVariable errordescr -ErrorAction SilentlyContinue  <### -ErrorAction SilentlyContinue  -> Update_Razvan: add this option at the end of each connection line after verifying that the function are ok #>
                        $Global:Domain = get-msoldomain -ErrorAction SilentlyContinue | Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"}  
                        $CurrentError = $errordescr.exception.message <### Update_Razvan: verify every error message on every connection endpoint and changed it accordingly: $errordescr | fl * -Force #>
                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($null -ne $Global:Error))
                &$Global:DisplayConnect
    }

    # Connect to Exchange Online PowerShell
    "EXO"  {    
                $Global:Error.Clear();

                # The loop for re-entering credentials in case they are wrong and for re-connecting
                $CurrentProperty = "Connect EXO"
                
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Exchange Online PowerShell"
                        $try++
                        
                        Write-Host "`nDo you connect to Exchange Online with MFA?" -ForegroundColor Cyan
                        $Global:mfa = get-choice "Yes", "No"
                        if ($mfa -eq "y")
                        {
                            "MFA"
                            Write-Host "When you use the Exchange Online Remote PowerShell Module, your session will end after one hour, which can be problematic for long-running scripts or processes. To avoid this issue, use Trusted IPs to bypass MFA for connections from your intranet." -ForegroundColor Yellow
                            if ($null -eq ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | `
                                Where-Object { $_ -notmatch "_none_" })) 
                            {
                                Write-Host "You requested to connect to Exchange Online with MFA but you don't have Exchange Online Remote PowerShell Module installed" -ForegroundColor Red
                                Write-Host "Please check the article https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps" -ForegroundColor Red
                                Read-Host "The script will now exit. Press any key then [Enter] to exit"
                                Disconnect-All
                                Exit
                            }
                            Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | `
                                Where-Object { $_ -notmatch "_none_" } | Select-Object -First 1)
                            $EXOSession = New-ExoPSSession  -PSSessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop
                            $CurrentError = $errordescr.exception 
                            Import-Module (Import-PSSession $EXOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue
                            $CurrentDescription = "Success"
                        }
                        else 
                        {
                            "Without MFA"
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
                        }
                        write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription

                        #Creating EXO PS Session
                        #$Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue | Out-Null
                        #$CurrentError = $errordescr.exception  
                        #Import-Module (Import-PSSession $EXOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue | out-null

                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($Global:Error)) 
                &$Global:DisplayConnect
    }

    # Connect to EXO2
    "EXO2"  {
        $Global:Error.Clear();
        If ($null -eq $Global:O365Cred) {
                &$Global:UserCredential
        }
        # The loop for re-entering credentials in case they are wrong and for re-connecting
        
        $CurrentProperty = "Connect EXO2"
        Do {
                # Defining the banner variable and clear the errors
                $Global:Error.Clear();
                $Global:banner = "Exchange Online v2 PowerShell"
                $try++
                $CurrentError = $errordescr.exception
                Import-Module ExchangeOnlineManagement -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                Connect-ExchangeOnline -Credential $global:O365Cred -PSSessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop 
                # Connection Errors check (mostly for wrong credentials reasons)
                &$Global:CredentialValidation
    }
    while (($Try -le 2) -and ($Global:Error)) 
    
    &$Global:DisplayConnect
}

    # Connect to EOP
    "EOP"  {
                $Global:Error.Clear();
                If ($null -eq $Global:O365Cred) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                
                $CurrentProperty = "Connect EOP"
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Exchange Online Protection PowerShell"
                        $try++
                        # Creating EOP PS session
                        $Global:EOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.protection.outlook.com/powershell-liveid" -Credential $global:O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $EOPSession  -AllowClobber -DisableNameChecking ) -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                        $Global:Domain = Get-AcceptedDomain | Where-Object {$_.name -like "*.onmicrosoft.com" } | Where-Object {$_.name -notlike "*mail.onmicrosoft.com"} 
            }
            while (($Try -le 2) -and ($Global:Error)) 
            
            &$Global:DisplayConnect
    }

    # Connect to Compliance Center Online
    "SCC"  {
                $Global:Error.Clear();
                If ($null -eq $Global:O365Cred) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                
                $CurrentProperty = "Connect SCC"
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Security and Compliance Center Powershell"
                        $try++ 
                        $Global:SCCSession = New-PSSession -ConfigurationName Microsoft.Compliance -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" -Credential $global:O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $SCCSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -Prefix CC -ErrorAction SilentlyContinue
                        #Credentials check
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($null -ne $Global:Error)) 
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
               If ($O365Cred.UserName -like "*.onmicrosoft.com")
                            {
                            If ($O365Cred.UserName -like "*.mail.onmicrosoft.com")
                                    {
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
                 $Global:Error.Clear();
                 If ($null -eq $Global:O365Cred) {
                        &$Global:UserCredential
                 }
                 # The loop for re-entering credentials in case they are wrong and for re-connecting
                 $CurrentProperty = "Connect AIPService"
                 
                 Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "AIPService PowerShell"
                        $try++
                        # Import AIPService module
                        Import-Module AIPService
                        # Creating a new AIPService PS Session
                        Connect-AIPService -Credential $global:O365Cred -ErrorVariable errordescr 
                        $CurrentError = $errordescr.exception
                        # Credentials check
                        &$Global:CredentialValidation
                    }
                 while (($Try -le 2) -and ($null -ne $Global:Error)) 
                 &$Global:DisplayConnect
    }
  }
#endregion Connection scripts region
}

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


function Disconnect-All {
    
    $CurrentDescription = "Disconnect is successful!"

    try {
            # Check and remove EXO session
            if($Global:O365Session){
                Remove-PSSession $Global:EXOSession}
                            
            # Check and remove EOP session
            if($Global:EOPSession){
                Remove-PSSession $Global:EOPSession}

            # Check and remove SCO session
            if($Global:SCCSession){
                Remove-PSSession $Global:SCOSession}
            
            # Check and remove S4B session
            if($global:sfboSession){
                Remove-PSSession $global:sfboSession}
            
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
    }



Function    Start-O365Troubleshooters
{
    Set-GlobalVariables
    Start-O365TroubleshootersMenu
}

Function Start-O365TroubleshootersMenu {
    $menu=@"
    1  Encryption: Office Message Encryption General Troubleshooting
    2  Security: Analyze compromise account/tenant
    3  Mail Flow: SMTP Relay Test
    4  Tools: Exchange Online Audit Search
    5  Tools: Unified Logging Audit Search
    6  Tools: Azure AD Audit Log Search
    7  Tools: Find all users with a specific RBAC Role
    8  Tools: Find all users with a specific RBAC Role
    9  Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox
    10 Tools: Decode SafeLinks URL
    Q  Quit
     
    Select a task by number or Q to quit
"@

Write-Host "My Menu" -ForegroundColor Cyan
$r = Read-Host $menu

Switch ($r) {
    "1" {
        Write-Host "Action Plan: Office Message Encryption General Troubleshooting" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-OfficeMessageEncryption.ps1
    }
     
    "2" {
        Write-Host "Action Plan: Analyze compromise account/tenant" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-CompromisedInvestigation.ps1
    }
     
    "3" {
        Write-Host "Action Plan: SMTP Relay Test" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-Office365Relay.ps1
    }
    "4" {
        Write-Host "Tools: Exchange Online Audit Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-ExchangeOnlineAuditSearch.ps1
    }
    "5" {
        Write-Host "Tools: Unified Logging Audit Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-UnifiedAuditLogSearch.ps1
    }
    "6" {
        Write-Host "Tools: Azure AD Audit Log Search" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-AzureADAuditLogSearch.ps1
    }
    "7" {
        Write-Host "Tools: Find all users with a specific RBAC Role" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-FindUserWithSpecificRbacRole.ps1
    }
    "8" {
        Write-Host "Tools: Find all users with all RBAC Role" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-AllUsersWithAllRoles.ps1
    }
    
    "9" {
        Write-Host "Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-MailboxDiagnosticLogs.ps1
    }
     
    "10" {
        Write-Host "Tools: Decode SafeLinks URL" -ForegroundColor Green
        . $script:modulePath\ActionPlans\Start-DecodeSafeLinksURL.ps1
    }

    "Q" {
        Write-Host "Quitting" -ForegroundColor Green
        Start-Sleep -Seconds 2
        Disconnect-all 
        exit
    }
     
    default {
        Write-Host "I don't understand what you want to do. Will reload the menu!" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Clear-Host
        Start-O365TroubleshootersMenu 
     }
    } 


}
