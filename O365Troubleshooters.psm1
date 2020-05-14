

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
    param ([ValidateSet("msol","exo","eop","sco","spo","sfb","aadrm")][Parameter(Mandatory=$true)] 
            $O365Service 
    )
    $Try = 0
    $global:errordesc = $null
    $Global:O365Cred=$null
    
#region Module Checks
    # $O365Service = "MSOL", "EXO" - Checking if the Azure Active Directory Module for Windows PowerShell (64-bit version) modules are installed on the machine
    If ( $O365Service -match "MSOL") {
            If ((get-module -ListAvailable -Name *MSOnline*).count -ne 2 ) {
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
            If ((Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell).count -ne 1) {
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
            If ((Get-Module -ListAvailable -Name LyncOnlineConnector).count -ne 1) {
                $CurrentProperty = "CheckingSFB Module"
                Write-Host "`nSkype for Business Online, Windows PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=39366' in order to download and install the module" -ForegroundColor Yellow
                $CurrentDescription = "Skype for Business Online, Windows PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=39366' in order to download and install the module"
                Write-Host "`nNow the script will stop." -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                Read-Host
                Exit
            } 
    }

    If ( $O365Service -match "AADRM") {
            If ((Get-Module -ListAvailable -Name AADRM).count -ne 1) {
                $CurrentProperty = "CheckingAADRM Module"
                Write-Host "`nAzure Active Directory Right Management PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=30339' in order to download and install the module" -ForegroundColor Yellow
                $CurrentDescription = "Azure Active Directory Right Management PowerShell module is not installed. Please access 'https://www.microsoft.com/en-us/download/details.aspx?id=30339' in order to download and install the module"
                Write-Host "`nNow the script will stop." -ForegroundColor Red
                write-log -Function "Connect-O365PS" -Step $CurrentProperty -Description $CurrentDescription
                Read-Host
                Exit
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
                If ($Global:O365Cred -eq $null) {
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
                        Connect-MsolService �Credential $O365Cred -ErrorVariable errordescr -ErrorAction SilentlyContinue  <### -ErrorAction SilentlyContinue  -> Update_Razvan: add this option at the end of each connection line after verifying that the function are ok #>
                        $Global:Domain = get-msoldomain -ErrorAction SilentlyContinue | ?{$_.name -like "*.onmicrosoft.com" } | ? {$_.name -notlike "*mail.onmicrosoft.com"}  
                        $CurrentError = $errordescr.exception.message <### Update_Razvan: verify every error message on every connection endpoint and changed it accordingly: $errordescr | fl * -Force #>
                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($Global:Error -ne $null))
                &$Global:DisplayConnect
    }

    # Connect to Exchange Online PowerShell
    "EXO"  {    
                $Global:Error.Clear();
                If ($Global:O365Cred -eq $null) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                $CurrentProperty = "Connect EXO"
                
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Exchange Online POwerShell"
                        $try++
                        #TRy & Catch

                        try {

                        $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction Stop 
                        $CurrentError = $errordescr.exception  
                        Import-Module (Import-PSSession $EXOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -ErrorAction SilentlyContinue
                        $CurrentDescription = "Success"
                        $Global:Domain = Get-AcceptedDomain | ?{$_.name -like "*.onmicrosoft.com" } | ? {$_.name -notlike "*mail.onmicrosoft.com"}  
       
               }
      catch {
              
        $CurrentDescription = "`""+$CurrentError.ErrorRecord.Exception +"`""

               
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

    # Connect to EOP
    "EOP"  {
                $Global:Error.Clear();
                If ($Global:O365Cred -eq $null) {
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
                        $Global:EOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.protection.outlook.com/powershell-liveid" -Credential $O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue 
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $EOPSession  -AllowClobber -DisableNameChecking ) -Global -DisableNameChecking  -ErrorAction SilentlyContinue
                        # Connection Errors check (mostly for wrong credentials reasons)
                        &$Global:CredentialValidation
                        $Global:Domain = Get-AcceptedDomain | ?{$_.name -like "*.onmicrosoft.com" } | ? {$_.name -notlike "*mail.onmicrosoft.com"} 
            }
            while (($Try -le 2) -and ($Global:Error)) 
            
            &$Global:DisplayConnect
    }

    # Connect to Compliance Center Online
    "SCO"  {
                $Global:Error.Clear();
                If ($Global:O365Cred -eq $null) {
                        &$Global:UserCredential
                }
                # The loop for re-entering credentials in case they are wrong and for re-connecting
                
                $CurrentProperty = "Connect SCO"
                Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Security and Compliance Center Powershell"
                        $try++ 
                        $Global:SCOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" -Credential $O365Cred -Authentication "Basic" -AllowRedirection -SessionOption $PSsettings -ErrorVariable errordescr -ErrorAction SilentlyContinue
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $SCOSession  -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -Prefix SCO -ErrorAction SilentlyContinue
                        #Credentials check
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($Global:Error -ne $null)) 
                &$Global:DisplayConnect
    }
    
    #Connect to SharePoint Online PowerShell
    "SPO" {
                $Global:Error.Clear();
                Import-Module MSOnline ;
                If ($Global:O365Cred -eq $null) {
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
               

                If ($domain -ne $null) {
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
                while (($Try -le 2) -and ($Global:Error -ne $null)) 
                &$Global:DisplayConnect                     
    }

    # Connect to Skype Online PowerShell
    "SFB" {
                $Global:Error.Clear();
                Import-Module MSOnline ;
                If ($Global:O365Cred -eq $null) {
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
                        $global:sfboSession = New-CsOnlineSession -Credential $O365Cred -ErrorVariable errordescr
                        $CurrentError = $errordescr.exception
                        Import-Module (Import-PSSession $sfboSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking 
                        # Credentials check
                        &$Global:CredentialValidation
                }
                while (($Try -le 2) -and ($Global:Error -ne $null)) 
                &$Global:DisplayConnect
    }

    # Connect to AADRM Service PowerShell
    "AADRM" {
                 $Global:Error.Clear();
                 If ($Global:O365Cred -eq $null) {
                        &$Global:UserCredential
                 }
                 # The loop for re-entering credentials in case they are wrong and for re-connecting
                 $CurrentProperty = "Connect AADRM"
                 
                 Do {
                        # Defining the banner variable and clear the errors
                        $Global:Error.Clear();
                        $Global:banner = "Azure AD Right Management Online PowerShell"
                        $try++
                        # Import AADRM module
                        Import-Module AADRM
                        # Creating a new AADRM PS Session
                        Connect-AadrmService -Credential $O365Cred -ErrorVariable errordescr
                        $CurrentError = $errordescr.exception
                        # Credentials check
                        &$Global:CredentialValidation
                    }
                 while (($Try -le 2) -and ($Global:Error -ne $null)) 
                 &$Global:DisplayConnects
    }
  }
#endregion Connection scripts region
}



Function Set-GlobalVariables {
 Write-Host 
 $global:FormatEnumerationLimit = -1
 $global:ts = Get-Date -Format yyyyMMdd_HHmmss
 $global:Path =[Environment]::GetFolderPath("Desktop")
 $Global:Path += "\PowerShellOutputs"
 $global:WSPath = "$Path\PowerShellOutputs_$ts"
 $global:starline = New-Object String '*',5
 #$Global:ExtractXML_XML = "Get-MigrationUserStatistics ", "Get-ImapSubscription "
 

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
 
}


#???
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
        foreach {

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
     return $Global:Option
    
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
               $version = Get-ItemProperty $v4Directory -name Version | select -expand Version
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
    $op_ver = Get-WmiObject Win32_OperatingSystem | select Caption, OSArchitecture
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


function disconnect-all {
    
    $CurrentDescription = "Disconnect is successful!"

    try {
            # Check and remove EXO session
            if($Global:O365Session){
                Remove-PSSession $Global:O365Session}
            
            # Check and remove EOP session
            if($Global:EOPSession){
                Remove-PSSession $Global:EOPSession}

            # Check and remove SCO session
            if($Global:SCOSession){
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
    Read-Host -Prompt "Please press [Enter] to exit"
    }

Function Start-AP_OfficeMessageEncryption {
    # Required function to set Global Variables
    Set-GlobalVariables
        
    # Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
    $Workloads = "exo", "sco", "aadrm"
    Connect-O365PS $Workloads
        
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
        
    # Main Function
        
    # Disconnecting
    disconnect-all  
}

Function Start-O365Troubleshooters {
    $menu=@"
    1 Office Message Encryption General Troubleshooting
    2 Analyze compromise account/tenant
    3 SMTP Relay Test
    4 Tools: Exchange Online Audit Search
    5 Tools: Unified Logging Audit Search
    6 Tools: Find all users with a specific RBAC Role
    7 Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox
    Q Quit
     
    Select a task by number or Q to quit
"@

Write-Host "My Menu" -ForegroundColor Cyan
$r = Read-Host $menu

Switch ($r) {
    "1" {
        Write-Host "Running: Office Message Encryption General Troubleshooting" -ForegroundColor Green
        Start-AP_OfficeMessageEncryption 
    }
     
    "2" {
        Write-Host "Analyze compromise account/tenant" -ForegroundColor Green
        #insert your code here
    }
     
    "3" {
        Write-Host "SMTP Relay Test" -ForegroundColor Green
        Start-AP_Office365Relay
    }
    "4" {
        Write-Host "Tools: Exchange Online Audit Search" -ForegroundColor Green
        #insert your code here
    }
    "5" {
        Write-Host "Tools: Unified Logging Audit Search" -ForegroundColor Green
        #insert your code here
    }
    "6" {
        Write-Host "Find all users with a specific RBAC Role" -ForegroundColor Green
        #insert your code here
    }
    "7" {
        Write-Host "Tools: Export All Available  Mailbox Diagnostic Logs for a given mailbox" -ForegroundColor Green
        Start-AP_MailboxDiagnosticLogs
    }
     
    "Q" {
        Write-Host "Quitting" -ForegroundColor Green
        exit
    }
     
    default {
        Write-Host "I don't understand what you want to do. Will reload the menu!" -ForegroundColor Yellow
        Start-O365Troubleshooters 
     }
    } 


}

Function Start-AP_MailboxDiagnosticLogs {
    
    # Required function to set Global Variables
    Set-GlobalVariables
        
    # Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
    $Workloads = "exo"
    Connect-O365PS $Workloads
        
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
        
    # Main Function
    $Ts= get-date -Format yyyyMMdd_HHmmss

    Write-Host "`nPlease input the path were the files will be saved" -ForegroundColor Green
    $ExportPath = Read-Host

    if ($ExportPath[-1] -eq "\") {
        $ExportPath = $ExportPath.Substring(0,$ExportPath.Length-1)
    }

    If (Test-Path -Path $ExportPath) {
        #Write-Host "`nThe path exist!" -ForegroundColor Green
    }
    else {
        Write-Host "`nThe output folder doesn't exist or is not valid! Please create or use an existing one and re-run the script. Press [Enter] to exit" -ForegroundColor Red
        Read-Host
        Exit
    }

    #endregion

    #region MbxDiagLogs

    Write-Host "`nPlease input the mailbox for which you want to see MailboxDiagnosticLogs: " -ForegroundColor Green
    $mbx = Read-Host

    # Check if mailbox exist
    $previousErrorActionPreference = $global:ErrorActionPreference
    $global:ErrorActionPreference = 'Stop'
    try{
        Get-Mailbox $mbx | Out-Null
    }
    Catch{
        Write-Host "`nThe mailbox $mbx doesn't exist. Press [Enter] to exit"
        Read-Host
        $ErrorActionPreference = $previousErrorActionPreference
        Exit
    }
    $global:ErrorActionPreference = $previousErrorActionPreference


    # Getting available components that can be exported 
    $previousErrorActionPreference = $global:ErrorActionPreference
    $global:ErrorActionPreference = 'Stop'
    $global:error.Clear()
    Try {
        Export-MailboxDiagnosticLogs $mbx -ComponentName TEST 
    }
    Catch {
        #Write-Host "in catch"
        $global:MbxDiagLogs = ((($global:error[0].Exception.Message -Split "Available logs: ")[1] -replace "'") -split ",") -replace " "
    }

    $global:ErrorActionPreference = $previousErrorActionPreference

     # Export-MailboxDiagnosticLogs with ComponentName
    $option = ( $global:MbxDiagLogs + "ALL")|Out-GridView -PassThru -Title "Choose a specific ComponentName or the last one for ALL"
    if ($option -ne "ALL") {
        Write-Host "`nGetting $option logs" -ForegroundColor Yellow 
        $option | ForEach-Object {
            Export-MailboxDiagnosticLogs $mbx -ComponentName  $_ | Tee-Object $ExportPath\$($Ts)_$_.txt
        } 
    }
    else {
        $MbxDiagLogs |ForEach-Object{
            Write-Host "`nGetting $_ logs" -ForegroundColor Yellow 
            Export-MailboxDiagnosticLogs $mbx -ComponentName  $_ | Tee-Object $ExportPath\$($Ts)_$_.txt
        }
    }


    # Export-MailboxDiagnosticLogs with ExtendedProperties
    Write-Host "You can view & filter ExtendedProperties in the Grid View window." -ForegroundColor Yellow
    $extendLogs = Export-MailboxDiagnosticLogs $mbx -ExtendedProperties
    $ExtendedProps = [XML]$extendLogs.MailboxLog
    $ExtendedProps.Properties.MailboxTable.Property | Select-Object name,value | Out-GridView -Title "All ExtendedProperties with values (you can filter here to find what is interesting for you; e.g: use `"ELC`" for MRM properties)"
    $ExtendedProps.Properties.MailboxTable.Property | Select-Object name,value |Out-File $ExportPath\$($Ts)_ExtendedProperties.txt

    Write-Host "`nOutput was exported in the following location: $ExportPath" -ForegroundColor Yellow 

     # Disconnecting
     disconnect-all  

}

#region Office365RelayDependencies
<#
Office 365 Relay Script Simulates a local application attempting to perform one of "SMTP Client Submission"/"SMTP Relay"/"Direct Send".
###leverages Send-MailMessage to handle the email submission.
###handles validation for Input data and error handling.
###has a central repository for known errors for this submission type and provides troubleshooting suggestions.

Aside from offering troubleshooting opportunities, it also grants option for customizing Email Body and Attachment.
Some built in features that can be used for AntiSpam testing:
###Standard test message body with timestamp
###Gtube test
###SpamLink SafeLink URL
###Input from Console
###*.htm / HTML formatted EML from Desktop
#>

<#
Validates Email Address via Regex
#>

function Get-ValidEmailAddress([string]$EmailAddressType)
{
    [string]$EmailAddress = Read-Host "Enter Valid $EmailAddressType"

    if($EmailAddress -match "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$")
    {
        return $EmailAddress
    }
    else
    {
        Get-ValidEmailAddress($EmailAddressType)
    }
}

<#
Validates Domain via Regex
Includes a Switch purely for future extention of the script
#>
function Get-ValidDomain([string]$DomainType)
{
    [string]$Domain = Read-Host "Enter Valid $DomainType Domain Name"

    switch($DomainType)
    {
        "Initial *.onmicrosoft.com"
        {
            if($Domain -match "^[A-Z0-9]+.onmicrosoft.com$")
            {
                return $Domain
            }
            else
            {
                Get-ValidDomain($DomainType)
            }
        }

        default
        {
        Write-Host "Unknown Domain Type Input Received"
        Get-RetryMenu
        }

    }
}


<#
Collects a List of Valid Recipient Email Addresses:
#Ask for Recipients
#Provide done to end list
#>

function Get-Office365RelayRecipients()
{   
    $Office365RelayErrorList.Clear()
    [int]$Office365RelayRecipientCount = Read-Host "Enter a Number of Recipients" -ErrorAction SilentlyContinue `
                                            -ErrorVariable +Office365RelayErrorList
    
    if( ($null -eq $Office365RelayErrorList[0]) -and ($Office365RelayRecipientCount -gt 0 ) -and ($Office365RelayRecipientCount -le 500 ) )
    {
        [string]$EmailAddressType = "RcptTo Email Address"

        [int]$i = 0
        
        while($i -lt $Office365RelayRecipientCount)
        {
            [string[]]$Office365RelayRecipients += Get-ValidEmailAddress([string]$EmailAddressType)
            $i++
        }

        return $Office365RelayRecipients
    }

    else 
    {
        Get-Office365RelayRecipients
    }
}

<#
Collect cloud smarthost for Smtp Relay/Direct Send function, performs:
    -Collects Cloud Smarthost
    -checks Resolve-DnsName to confirm that A record is correctly propagated for it.
#>
function Find-O365RelaySmarthost()
{
    $Office365RelayErrorList.Clear()

    [string]$DomainType = "Initial *.onmicrosoft.com"

    $InitialDomain = [string](Get-ValidDomain($DomainType))

    $MXResolution = Resolve-DnsName -Type MX -Name $InitialDomain -ErrorAction Continue -ErrorVariable +Office365RelayErrorList

    if ($MXResolution.Type -eq "MX")
    {   
        $Office365RelaySmarthost = Resolve-DnsName -Type A -Name $MXResolution.NameExchange -ErrorAction Continue `
                                    -ErrorVariable +Office365RelayErrorList

        if ($null -eq $Office365RelayErrorList[0])
        {
            return $Office365RelaySmarthost[0].Name
        }

        else
        {
            Get-ActionPlan("SmarthostFunctionErrors")
        }
    }

    else 
    {
        Get-ActionPlan("SmarthostFunctionErrors")
    }
}

<#
Collect Credentials function, performs :
    -clean-up for email address format based on Regex.
    -runs Powershell Connection check to validate credentials.
    -closes Powershell connection after confirming login works.
#>
function Get-AuthenticationCredentials()
{
    $Office365RelayErrorList.Clear()

    $O365SenderCred = Get-Credential

    $O365AuthenticationCredentialsSession = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365SenderCred `
        -Authentication Basic -AllowRedirection -ErrorAction Continue -ErrorVariable +Office365RelayErrorList
        
    if ($null -eq $Office365RelayErrorList[0])
    {
        Remove-PSSession -InstanceId $O365AuthenticationCredentialsSession.InstanceId
        return $O365SenderCred
    }

    else
    {
        Get-ActionPlan("AuthenticationFunctionErrors")
    }
}

<#
Test-Office365RelayScriptItemPath checks for item path
returns true or false (bool).
#>
function Test-Office365RelayScriptItemPath([string] $Office365RelayScriptFilePath)
{
    return Test-Path $Office365RelayScriptFilePath
}

<#
Get-MessageBody function provides:
-standard test body with timestamp
-Gtube test
-SpamLink SafeLink URL
-option to input from Console
-Custom content from htm files
#>
function Get-MessageBody([DateTime] $d)
{
    [string] $MessageBodyType = Read-Host "`r`nSelect a type of Message Body`r`n
A : Standard test message body with timestamp
B : Gtube test
C : SpamLink SafeLink URL
D : Input from Console
E : *.htm / HTML formatted EML from Desktop

Answer"

    switch($MessageBodyType.ToUpper())
    {
        A
        {
            return "This is the body of test message. The message was sent at: $d"
        }

        B
        {
            return "XJS*C4JDBQADN1.NSBN3*2IDNEN*GTUBE-STANDARD-ANTI-UBE-TEST-EMAIL*C.34X"
        }

        C
        {
            return "https://www.spamlink.contoso.com"
        }

        D
        {
            return (Read-Host "Enter a Message Body and press enter")
        }

        E
        {
            $Path=[Environment]::GetFolderPath("Desktop")
            [string] $MessageBodyFile = Read-Host "Enter the target EML file full name with extension
for example : EmlFile.eml
Note: Only eml file format can be parsed`r`n
Full File Name"
            [bool]$isPathValid = Test-Office365RelayScriptItemPath([string] "$Path\$MessageBodyFile")
            if($isPathValid)
            {   
                $emlContent = Get-Content "$Path\$MessageBodyFile" -Encoding utf8
                [int] $startOfEmlContent = ($emlContent | Select-String '<!DOCTYPE html>').LineNumber - 1
                [int] $endOfEmlContent = ($emlContent | Select-String '</html>').LineNumber - 1

                if ($null -ne $emlContent)
                {
                return $emlContent[$startOfEmlContent..$endOfEmlContent]
                }

                else 
                {
                Write-Host "`r`nNo Content Found in the pointed file`r`n" -ForegroundColor Red
                Get-MessageBody($d)
                }
            }

            else 
            {
                Write-Host "Content file not found, check path and file name are valid"
                Get-MessageBody($d)
            }
        }

        default
        {Get-MessageBody($d)}
    }  
}

<#
Get-MessageAttachment function provides:
-Eicar test
-Custom attachment from local machine
#>
function Get-MessageAttachment()
{
    [string] $MessageAttachmentType = Read-Host "`r`nSelect a type of Attachment`r`n
    A : File from Desktop
    B : No Attachment
    
    Answer"
    
        switch($MessageAttachmentType.ToUpper())
        {
            A
            {
                $Path=[Environment]::GetFolderPath("Desktop")
                [string] $MessageAttachmentFile = Read-Host "Enter the target htm file full name with extension
for example : attachmentfile.csv`r`n
Answer"         
                $isPathValid = Test-Office365RelayScriptItemPath([string] "$Path\$MessageAttachmentFile")
                if($isPathValid)
                {   
                    return "$Path\$MessageAttachmentFile"
                }

                else 
                {
                    Write-Host "Attachment file not found, check path and file name are valid"
                    Get-MessageAttachment
                }
            }

            B
            {
                return "noAttachment"
            }

            default
            {
                Get-MessageAttachment
            }
        }
}

<#
SMTP Client Submission function:
-sends SMTP Client Submission email
-tries to identify known error 
-writes DSN in logs.
#>
function Send-ClientSubmission([PSCredential]$Credentials)
{
    $Office365RelayErrorList.Clear()

    [int]$Port = Read-Host "Input open outbound port (25 or 587)" -ErrorAction SilentlyContinue -ErrorVariable +Office365RelayErrorList

    if ( ($null -eq $Office365RelayErrorList[0]) -and (($Port -eq 25) -or ($Port -eq 587)) )
    {
        $d = Get-Date
        [string] $o365RelayMessageBody = Get-MessageBody($d)
        [string] $o365RelayMessageAttachment = Get-MessageAttachment

        write-host "Sending Message..."
        
        switch -wildcard ($o365RelayMessageAttachment)
        { 
            'noAttachment'
            {         
                Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver smtp.office365.com -subject "SMTP Client Submission Email - $d" `
                                    -body $o365RelayMessageBody -Credential $Credentials -UseSsl -BodyAsHtml -port $Port `
                                    -ErrorAction Continue -WarningAction SilentlyContinue -ErrorVariable +Office365RelayErrorList
            }

            '*Users*Desktop*'
            {
                Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver smtp.office365.com -subject "SMTP Client Submission Email - $d" `
                                    -body $o365RelayMessageBody -Credential $Credentials -Attachments $o365RelayMessageAttachment `
                                    -UseSsl -BodyAsHtml -port $Port -ErrorAction Continue -WarningAction SilentlyContinue `
                                    -ErrorVariable +Office365RelayErrorList
            }
            
            Default 
            {
            Write-Host "Not Implemented yet, please go back to Main Menu or Exit"
            Get-RetryMenu
            }
        }
        

                          
        
        if ($null -ne $Office365RelayErrorList[0])
        {
            Get-ActionPlan("SmtpClientSubmissionFunctionErrors")                   
        }
    
        else
        {
            Write-Host "`r`nEmail Sent Succesfully"
        
            Get-RetryMenu
        }
    }

    else
    {
        Send-ClientSubmission($Credentials)
    }

    
}

<#
SMTP Relay function:
-sends email via Anonymous SMTP Submission
-tries to identify error issue 
-writes DSN in logs.
#>
function Send-SMTPRelay([string] $O365RelaySmarthost)
{
    $Office365RelayErrorList.Clear()

    [int]$Port = 25

    if ($null -eq $Office365RelayErrorList[0])
    {
        write-host "Sending Message..."
        $d = Get-Date
        Send-mailmessage -to $Office365RelayRecipients -from $O365SendAs -smtpserver $O365RelaySmarthost -subject "SMTP Relay Email - $d" `
                                -body "This is the body of test message. The message was sent at: $d" -UseSsl -BodyAsHtml -Port $Port `
                                -ErrorAction Continue -WarningAction SilentlyContinue -ErrorVariable +Office365RelayErrorList
                          
        if($null -ne $Office365RelayErrorList[0])
        {
            Get-ActionPlan("SmtpRelayFunctionErrors")            
        }
        
        else
        {
            Write-Host "`r`nEmail Sent Succesfully"

            Get-RetryMenu
        }
    }
}

<#
Error + Action Plan Central repository
#>
function Write-ScriptLog([string] $ErrorType) 
{
    $d = Get-Date
    $TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    "`r`n$FailedAction at $d $TimeZone generated Error:`r`n" + $Office365RelayErrorList | Out-File -Append $path\Logs\$ErrorType.txt
}
function Get-ActionPlan([string]$ErrorType)
{
    switch($ErrorType)
    {
        "SmtpRelayFunctionErrors"
        {
            [string] $FailedAction = "Email Sent As : $O365SendAs"
            Write-ScriptLog($ErrorType)

            switch -Wildcard ($Office365RelayErrorList[0].Exception.Message)
            {
                '*Unable to connect to the remote server*'
                {
                    Write-Host "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
`r`nFor instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
 
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                    Get-RetryMenu
                }
                
                '*connected party did not properly respond after a period of time*'
                {
                    Write-Host "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
`r`nFor instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
 
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                    Get-RetryMenu
                }

                '*5.7.606 Access denied, banned sending IP*'
                {
                    Write-Host "Public Outbound IP listed on : https://sender.office.com/ , visit
and follow directions to request removal.
For more information, you can visit : http://go.microsoft.com/fwlink/?LinkID=526655`r`n
Once the delist process is completed, allow 24 hours for full delist approval and propagation.
Note: We do not provide IP Statistics to support the reason for the block."
                    Get-RetryMenu
                }

                '*5.7.64 TenantAttribution; Relay Access Denied*'
                {
                    Write-Host "Relay was not allowed, If the intention was to perform SMTP Relay via
Certificate Scoped Inbound OnPrem Connector, this error is expected.
This script cannot use local certificates for mail submission.`r`n
If you are using an IP Scoped Inbound OnPrem Connector, the email 
is not attributed to your tenant as expected.`r`n
If you are using an IP Scoped Inbound OnPrem Connector and
SMTP Relay previously worked from this machine`r`n
Check the Outbound IP of the machine
You must install the Telnet client software before you can run this command. For more information, see 
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)`r`n
Run the tool from the command line by typing telnet
Type: open $O365RelaySmarthost 25
If you connected successfully to an Office 365 server, expect to receive a response line similar to this:
220 BN1BFFO11FD038.mail.protection.outlook.com Microsoft ESMTP MAIL Service ready at Mon, 18 Apr 2016 07:36:51 +0000`r`n
If the connection is not successful, then the network firewall or Internet Service Provider (ISP) may block port 25.
If the response does not contain ‘mail.protection.outlook.com Microsoft ESMTP MAIL Service’ check firewall configuration`r`n
Type the following command: EHLO FQDN.yourdomain.com, and then press Enter. You should receive the following response:
250-DB3FFO11FD036.mail.protection.outlook.com Hello [IP address]`r`n`r`n
-Ensure the Outbound IP mentioned in this reply is configured on the Inbound OnPrem Connector used for Relay.
https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-150/dn910993(v=exchg.150)?redirectedfrom=MSDN#using-the-ip-address-of-your-email-server`r`n
For general information on setting up Office 365 Relay Methods, check:
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/fix-issues-with-printers-scanners-and-lob-applications-that-send-email-using-off#emails-are-no-longer-being-sent-to-external-recipients`r`n
For further information regarding `"5.7.64 TenantAttribution; Relay Access Denied`", you can access: https://docs.microsoft.com/en-us/exchange/troubleshoot/connectors/office-365-notice"
                    Get-RetryMenu
                }

                '*5.7.1 Service unavailable, Client host*blocked using Spamhaus*'
                {
                    Write-Host "Follow instructions provided in the above listed DSN for delisting the IP.
This list is not managed by Microsoft and the Delist should be processed between the owner of the static outbound IP and Spamhaus."
                    Get-RetryMenu
                }

                '*4.4.62 Mail sent to the wrong Office 365 region. ATTR35*'
                {
                    Write-Host "If the intention was to perform SMTP Relay via Certificate Scoped Inbound OnPrem Connector, this error is expected. 
This script does not leverage your local certificate to handle the mail submission.`r`n
If you are using an IP Scoped Inbound OnPrem Connector, the email is not attributed to your tenant as expected.`r`n
If you are using an IP Scoped Inbound OnPrem Connector and
SMTP Relay previously worked from this machine`r`n
Check the Outbound IP of the machine
You must install the Telnet client software before you can run this command. For more information, see 
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)`r`n
Run the tool from the command line by typing telnet
Type: open $O365RelaySmarthost 25
If you connected successfully to an Office 365 server, expect to receive a response line similar to this:
220 BN1BFFO11FD038.mail.protection.outlook.com Microsoft ESMTP MAIL Service ready at Mon, 18 Apr 2016 07:36:51 +0000`r`n
If the connection is not successful, then the network firewall or Internet Service Provider (ISP) may block port 25.
If the response does not contain ‘mail.protection.outlook.com Microsoft ESMTP MAIL Service’ check firewall configuration`r`n
Type the following command: EHLO FQDN.yourdomain.com, and then press Enter. You should receive the following response:
250-DB3FFO11FD036.mail.protection.outlook.com Hello [IP address]`r`n`r`n
-Ensure the Outbound IP mentioned in this reply is configured on the Inbound OnPrem Connector used for Relay.
https://docs.microsoft.com/en-us/previous-versions/exchange-server/exchange-150/dn910993(v=exchg.150)?redirectedfrom=MSDN#using-the-ip-address-of-your-email-server`r`n
For general information on setting up Office 365 Relay Methods, check:
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/fix-issues-with-printers-scanners-and-lob-applications-that-send-email-using-off#emails-are-no-longer-being-sent-to-external-recipients`r`n
If you are using Direct Send, it seems you have reached an incorrect destination for the Office 365 Recipient.
Please check if you for incorrect DNS resolution for the recipient $O365RelaySmarthost MX`r`n
`r`nFor further information regarding this error, you can access:
https://docs.microsoft.com/en-us/exchange/troubleshoot/email-delivery/wrong-office-365-region-exo"
                    Get-RetryMenu
                }

                '*A socket operation was attempted to an unreachable network*'
                {
                    Write-Host "Check for local Internet Connectivity issues.
Try to perform the test using a different Internet Connection."
                    Get-RetryMenu
                }

                '*The operation has timed out*'
                {
                    Write-Host "Check for local Internet Connectivity issues.
Try to perform the test using a different Internet Connection."
                }
            
                default
                {
                    Write-Host "`r`nEmail Could not be sent. Error was not recognized, recorded in SMTPRelay errors.`r`n"

                    Get-RetryMenu
                }
            }
        }

        "SmtpClientSubmissionFunctionErrors"
        {   
            $ScriptLogAutenticatedUser = $Credentials.UserName
            [string] $FailedAction = "Email Sent As : $O365SendAs, authenticated as : $ScriptLogAutenticatedUser"
            Write-ScriptLog($ErrorType)
        
            switch -Wildcard ($Office365RelayErrorList[0].Exception.Message)
            {
                '*Unable to connect to the remote server*'
                {
                    Write-Host "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
`r`nFor instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019"
                    Get-RetryMenu
                }

                '*connected party did not properly respond after a period of time*'
                {
                    Write-Host "It seems the selected Port $Port is not open.
Test Port $Port via telnet and check for any local firewall/proxy restricting outbound traffic.
`r`nFor instructions on how to use telnet you can check the following article:
https://docs.microsoft.com/en-us/Exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019
 
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc771275(v=ws.10)"
                    Get-RetryMenu
                }

                '*SendAsDeniedException.MapiExceptionSendAsDenied*'
                {
                    Write-Host "The authenticated account does not have required SendAs permission for Sending Email Address.`r`n
Avoid using a single mailbox with Send As permissions for all your users. 
This method is not supported because of complexity and potential issues.
If you find yourself in this unsupported scenario, 
Direct Send / SMTP Relay should be used as alternatives.
https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3#compare-the-options"
                    Get-RetryMenu 
                }

                '*5.7.57 SMTP; Client was not authenticated to send anonymous mail*'
                {
                    Write-Host "Smtp Client Authentication may be disabled either for the account or on an organization level in EXO.
These settings can be checked via EXO Powershell :`r`n
Get-CasMailbox user@contoso.com|fl SmtpClientAuthenticationDisabled
Get-TransportConfig|fl SmtpClientAuthenticationDisabled`r`n
Note: To selectively enable authenticated SMTP for specific mailboxes only: 
    -disable authenticated SMTP at the organizational level ($true)
    -enable it for the specific mailboxes ($false)
    -and leave the rest of the mailboxes with their default value ($null).
`r`n
Be aware the Azure Security Defaults will automatically disable Legacy Authentication 
                                        and prevent SMTP Submission from being accepted
`r`n For more information on Azure Security defaults, check public article:
https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/concept-fundamentals-security-defaults"
                    Get-RetryMenu
                }

                default
                {
                    Write-Host "`r`nEmail Could not be sent.
Error was not recognized, recorded in Client Submission errors.`r`n"
                    Get-RetryMenu
                }

            }
        }

        "AuthenticationFunctionErrors"
        {   
            [string] $FailedAction = "Authentication with account "+$O365SenderCred.UserName
            Write-ScriptLog($ErrorType)
            
            Write-Host "`r`nAuthentication failed, please test authentication via browser
If this account is MFA enabled, please use an App Password instead of the Password to Authenticate.
You can create a new App Password as shown in the following article:
https://support.office.com/en-us/article/Create-an-app-password-for-Office-365-3e7c860f-bda4-4441-a618-b53953ee1183"

            Get-RetryMenu
        }

        "SmarthostFunctionErrors"
        {   
            [string] $FailedAction = "MX record lookup for smarthost "+$InitialDomain
            Write-ScriptLog($ErrorType)

            Write-Host "`r`nSmarthost lookup failed for $InitialDomain`r`n
Are the MX and corresponding A records propagated across internet when checking via 3rd party lookup tools online?
    -Yes : Investigate possible local DNS/Network issues.
    -No : Engage Microsoft Support to resolve the DNS record propagation."

            Get-RetryMenu
        }
    }
}

<#
Exit-ScriptAndSaveLogs:
-stops transcript
-displays log file location
-exits the script
#>
function Exit-ScriptAndSaveLogs() 
{
    [string] $logFileLocation = "`r`nAll logs have been saved to the following location: $path `r`n"
    Stop-Transcript
    Write-Host $logFileLocation -ForegroundColor Green
    Read-Host "Press Any Key to finalize Exit"
    Exit
}

<#
Retry Function will follow successful email submissions.
It will offer quit option with Stop Transcript. 
    -Rerun Get-MainMenu Menu
    -Stop Transcript
    -Display Log Path
    -Exit
#>
function Get-RetryMenu()
{
    $RetryInput = Read-Host "`r`nPress any key to reach Main Menu or Q to exit the script"
                    
    switch($RetryInput)
    {
        default {Get-MainMenu}
        Q       {Exit-ScriptAndSaveLogs}
    }
}


<#
Get-CashedOrFreshCredentials provides:
-option to reuse credentials from previous email submission
-option to input fresh credentials
#>
function Get-CashedOrFreshCredentials()
{
	if ($null -eq $Credentials)
    { return [PSCredential]$Credentials = Get-AuthenticationCredentials }
    
	else 
    {
		$ReUseCredentials = Read-Host "Do you want to use the same Credentials?
A : Yes
B : No
Answer"

		switch ($ReUseCredentials) 
		{
		    A { [PSCredential]$Credentials }
	        B { return Get-AuthenticationCredentials }
		    default { Get-CashedOrFreshCredentials }
		}
	}
}

<#
MainMenu function:
-reads RelayMethod input from Console
-writes input at Relay Method selection to a file
-calls required functions for each relay method (may be moved to each corresponding relay function instead)
-calls Exit-ScriptAndSaveLogs for quit from script
#>
function Get-MainMenu()
{
    Clear-Host
    
    $RelayMethod = Read-Host "Office 365 Relay Menu`r`n
A : Client Submission
B : SMTP Relay / Direct Send
Q : Quit Script`r`n

Answer"

    "RuntimeRelayMethodInput#$RuntimeRelayMethodCounter $RelayMethod"|Out-File -Append $path\Logs\ChoicesAtRuntime.txt
    $RuntimeChoiceCounter++

    switch($RelayMethod.ToUpper())
    {
        A 
        {   
            $Credentials = Get-CashedOrFreshCredentials
            [string]$O365SendAs = Get-ValidEmailAddress("From Email Address")
            #[string]$Recipients = Get-ValidEmailAddress("RcptTo Email Address")
            [string[]]$Office365RelayRecipients = Get-Office365RelayRecipients
            Send-ClientSubmission($Credentials)
        }
    
        B 
        {
            [string]$O365RelaySmarthost = Find-O365RelaySmarthost
            [string]$O365SendAs = Get-ValidEmailAddress("From Email Address")
            #[string]$Recipients = Get-ValidEmailAddress("RcptTo Email Address") 
            [string[]]$Office365RelayRecipients = Get-Office365RelayRecipients
            Send-SMTPRelay($O365RelaySmarthost)
        }
    
        Q
        {Exit-ScriptAndSaveLogs}
    
        default {Get-MainMenu}
    }
}
#endregion Office365RelayDependencies
function Start-AP_Office365Relay {
    $SendMailMessageDisclaimer ="
Warning : This Script is only recommended for testing purposes as it uses 'Send-MailMessage' cmdlet, 
which is currently considered obsolete.

This cmdlet does not guarantee secure connections to SMTP servers. While there is no immediate 
replacement available in PowerShell, we recommend you do not use Send-MailMessage at this time. 
See https://aka.ms/SendMailMessage for more information.`r`n"

    Write-Host $SendMailMessageDisclaimer -ForegroundColor Yellow

    Read-Host "`r`nPress any key to Continue, Ctrl+C to quit the script"

    Clear-Host
    
    $ts = Get-Date -Format yyyyMMdd_HHmm
    $path=[Environment]::GetFolderPath("Desktop")+"\$($ts)_RelayOptions"

    #Implement check if Log Folder already exists and provide alternative
    Write-Host "Created Directories on Desktop:"
    mkdir "$path"
    mkdir "$path\Logs"

    Write-Host "`r`n"

    Start-transcript -Path "$path\RelayTranscript_$ts.txt"

    Read-Host "`r`nPress any key to Continue, Ctrl+C to quit the script"

    $RuntimeChoiceCounter = 1
    $Office365RelayErrorList = @()
    Get-MainMenu
}