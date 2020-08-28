# MIT License
# 
# Copyright (c) 2020 O365Troubleshooters Team
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# Description:
    ### This script is analyzing the migration reports

# Author:
    ### Cristian Dimofte

# Versions:
    #####################################################################
    # Version   # Date          # Description                           #
    #####################################################################
    # 1.0       # 08/28/2020    # Initial script                        #
    #           #               #                                       #
    #####################################################################


##############################
# Common space for functions #
##############################
#region Functions


### LogsToAnalyze (Scope: Script) variable will contain mailbox migration logs for all affected users
[System.Collections.ArrayList]$script:LogsToAnalyze = @()

### ParsedLogs (Scope: Script) variable will contain parsed mailbox migration logs for all affected users
[System.Collections.ArrayList]$script:ParsedLogs = @()

### Get the timestamp from the time the scenario was accessed
[string]$ts = Get-Date -Format yyyyMMdd_HHmmss

### Create the location where to save logs related to MailboxMigration scenario
[string]$ExportPath = "$global:WSPath\MailboxMigration_$ts"
$null = New-Item -ItemType Directory -Path $ExportPath -Force


### <summary>
### Show-MailboxMigrationMenu function is used if the script is started without any parameters
### </summary>
function Show-MailboxMigrationMenu {

    $MailboxMigrationMenu=@"
    1  If you have the migration logs in an .xml file
    2  If you want to connect to Exchange Online in order to collect the logs
    B  Back to Action plans

    Select a task by number, or, B to go back to main menu: 
"@

    Write-Log -function "MailboxMigration - Show-MailboxMigrationMenu" -step "Loading the menu" -Description "Success"

    Clear-Host

    Write-Host $MailboxMigrationMenu -ForegroundColor White -NoNewline
    $SwitchFromKeyboard = Read-Host

    ### Providing a list of options
    Switch ($SwitchFromKeyboard) {

        ### If "1" is selected, the script will assume you have the mailbox migration logs in an .xml file
        "1" {
            Write-Log -function "MailboxMigration - Show-MailboxMigrationMenu" -step "Loading option 1" -Description "Success"
            Selected-FileOption
        }

        ### If "2" is selected, the script will connect you to Exchange Online
        "2" {
            Write-Log -function "MailboxMigration - Show-MailboxMigrationMenu" -step "Loading option 2" -Description "Success"
            Selected-ConnectToExchangeOnlineOption
        }

        ### If "B" is selected, move back to the "O365TroubleshootersMenu"
        "B" {
            Write-Log -function "MailboxMigration - Show-MailboxMigrationMenu" -step "Loading option `"B`"" -Description "Success"
            Start-O365TroubleshootersMenu
         }

        ### If you selected anything different than "1", "2" or "B", the Menu will reload
        default {
            Write-Host "You selected an option that is not present in the menu (Value inserted from keyboard: `"$SwitchFromKeyboard`")" -ForegroundColor Yellow
            Write-Host "Press any key to re-load the menu"
            Write-Log -function "MailboxMigration - Show-MailboxMigrationMenu" -step "Loading option `"default`"" -Description "Reload MailboxMigrationMenu"
            Read-Host
            Show-MailboxMigrationMenu
        }
    } 
}


### <summary>
### Selected-FileOption function is used when the information is already saved on a .xml file.
### </summary>
### <param name="FilePath">FilePath parameter is used when the script is started with the FilePath parameter.</param>
function Selected-FileOption {
    [CmdletBinding()]
    Param
    (        
        [string]$FilePath
    )

    [int]$TheNumberOfChecks = 1
    ### If FilePath was provided, the script will use it in order to validate if the information from this variable is a correct
    ### full path of an .xml file.
    if ($FilePath){
        try {
            ### The script validates that the path provided is of a valid .xml file.
            Write-Log -function "MailboxMigration - Selected-FileOption" -step "Start validation of `"$FilePath`" file" -Description "Success"
            [string]$PathOfXMLFile = Validate-XMLPath -XMLFilePath $FilePath
        }
        catch {
            ### In case of error, the script will ask to provide again the full path of the .xml file
            Write-Log -function "MailboxMigration - Selected-FileOption" -step "Ask for .xml path. Iteration $TheNumberOfChecks" -Description "Error validating initially provided .xml path"
            [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
        }
    }
    ### If no FilePath was provided, the script will ask to provide the full path of the .xml file
    else{
        Write-Log -function "MailboxMigration - Selected-FileOption" -step "Ask for .xml path. Iteration $TheNumberOfChecks" -Description "Success. No initial .xml path provided"
        [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
    }

    ### If PathOfXMLFile variable will match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed", we will stop the script
    if ($PathOfXMLFile -match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed") {
        Write-Log -function "MailboxMigration - Selected-FileOption" -step "Ask for .xml path. Iteration $TheNumberOfChecks" -Description "Error. $PathOfXMLFile matches `"NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed`""
        throw "The script will end, because the .xml file provided is not valid from PowerShell's perspective"
    }
    else {
        ### TheMigrationLogs variable will represent MigrationLogs collected using the Collect-MigrationLogs function.
        Write-Log -function "MailboxMigration - Selected-FileOption" -step "Start analyze of data from `"$FilePath`" file" -Description "Success"
        Collect-MigrationLogs -XMLFile $PathOfXMLFile
    }
}

### <summary>
### Validate-XMLPath function is used to check if the path provided is a valid .xml file.
### </summary>
### <param name="XMLFilePath">XMLFilePath parameter represents the path the script has to check if it is a valid .xml file.</param>
function Validate-XMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_})]
        [string]
        $XMLFilePath
    )

    ### Validating if the path has a length greater than 4, and if it is of an .xml file
    Write-Log -function "MailboxMigration - Validate-XMLPath" -step "Validating if `"$XMLFilePath`" is valid from PowerShell's perspective" -Description "Success"
    if (($XMLFilePath.Length -gt 4) -and ($XMLFilePath -like "*.xml")) {
        ### Validating if the .xml file was created by PowerShell
        $fileToCheck = new-object System.IO.StreamReader($XMLFilePath)
        if ($fileToCheck.ReadLine() -like "*http://schemas.microsoft.com/powershell*") {
            Write-Host
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " seems to be a valid .xml file. We will use it to continue the investigation." -ForegroundColor Green
            Write-Log -function "MailboxMigration - Validate-XMLPath" -step "`"$XMLFilePath`" is valid from PowerShell's perspective" -Description "Success"
        }
        ### If not, the script will set the XMLFilePath to NotAValidXMLFile. This will help in next checks, in order to start collecting the mailbox 
        ### migration logs using other methods
        else {
            Write-Log -function "MailboxMigration - Validate-XMLPath" -step "`"$XMLFilePath`" is not valid from PowerShell's perspective" -Description "We will set: XMLFilePath = `"NotAValidXMLFile`""
            $XMLFilePath = "NotAValidXMLFile"
        }

        $fileToCheck.Close()

    }
    ### If the path's length is not greater than 4 characters and the file is not an .xml file the script will set XMLFilePath to NotAValidPath.
    ### This will help in next checks, in order to start collecting the mailbox migration logs using other methods
    else {
        Write-Log -function "MailboxMigration - Validate-XMLPath" -step "`"$XMLFilePath`" is not valid from PowerShell's perspective" -Description "We will set: XMLFilePath = `"NotAValidPath`""
        $XMLFilePath = "NotAValidPath"
    }

    ### The script returns the value of XMLFilePath 
    return $XMLFilePath
}

### <summary>
### Ask-ForXMLPath function is used to ask for the full path of a .xml file.
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in order to do an 1-time effort to provide another path of the .xml file,
### in case first time when it was entered, there was a typo </param>
function Ask-ForXMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]$NumberOfChecks
    )

    [string]$PathOfXMLFile = ""
    if ($NumberOfChecks -eq "1") {
        ### Asking to provide the full path of the .xml file for the first time
        Write-Host
        Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "We are asking to provide the path of the .xml file. Iteration $NumberOfChecks" -Description "Success"
        Write-Host "Please provide the path of the .xml file: " -ForegroundColor Cyan
        Write-Host "`t" -NoNewline
        try {
            ### PathOfXMLFile variable will contain the full path of the .xml file, if it will be validated (it will be inserted from the keyboard)
            $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
        }
        catch {
            ### If error, the script is doing the 1-time effort to collect again the full path of the .xml file
            Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "Ask for the .xml path" -Description "Error when asked for a new .xml path. Retrying."
            $NumberOfChecks++
            $PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        ### The script is doing the 1-time effort to collect again the full path of the .xml file
        Write-Host
        Write-Log "[INFO] || Asking to provide the path of the .xml file again" -NonInteractive $true
        Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "Asking to provide the path of the .xml file again. Iteration $NumberOfChecks" -Description "Success"
        Write-Host "Would you like to provide the path of the .xml file again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes`t`t[N] No`t`t(default is `"N`"): " -NoNewline -ForegroundColor White
        $ReadFromKeyboard = Read-Host

        ### Checking if the path will be provided again, or no. If no, we will continue to collect the mailbox migration logs, using other methods.
        [bool]$TheKey = $false
        Switch ($ReadFromKeyboard) 
        { 
          Y {$TheKey=$true} 
          N {$TheKey=$false} 
          Default {$TheKey=$false} 
        }

        if ($TheKey) {
            ### If YES was selected, we are asking to provide the path of the .xml file again
            Write-Host
            Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "Please provide again the path of the .xml file" -Description "Success"
            Write-Host "Please provide again the path of the .xml file: " -ForegroundColor Cyan
            Write-Host "`t" -NoNewline
            try {
                ### Validating the path of the .xml file
                $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
            }
            catch {
                ### If error, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to stop the collection of the logs
                Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "Agreed to provide new .xml file. Unable to get an .xml file valid from PowerShell's perspective" -Description "We will set: PathOfXMLFile = `"ValidationOfFileFailed`""
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            ### If NO was selected, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
            Write-Log -function "MailboxMigration - Ask-ForXMLPath" -step "Agreed not to provide new .xml file" -Description "We will set: PathOfXMLFile = `"ValidationOfFileFailed`""
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }

    ### The function returns the full path of the .xml file, or ValidationOfFileFailed
    return $PathOfXMLFile
}


### <summary>
### Collect-MigrationLogs function is used to collect the mailbox migration logs
### </summary>
### <param name="XMLFile">XMLFile represents the .xml file from which we want to import the mailbox migration logs </param>
### <param name="ConnectToExchangeOnline">ConnectToExchangeOnline parameter will be used to connect to Exchange Online, and collect the 
### needed mailbox migration logs, based on the migration type used </param>
### <param name="ConnectToExchangeOnPremises">ConnectToExchangeOnPremises parameter will be used to connect to Exchange On-Premises, and collect the 
### the output of Get-MailboxStatistics (the MoveHistory part), for the affected user </param>
function Collect-MigrationLogs {
    [CmdletBinding()]
    Param (
        [parameter(Mandatory=$true,
        ParameterSetName="XMLFile")]
        [string]
        $XMLFile,
        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $true)]
        [switch]$ConnectToExchangeOnline,

        [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $true)]
        [switch]$ConnectToExchangeOnPremises,

        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
        [string[]]$AffectedUsers,

        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged")]
        [string]$MigrationType,

        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [string]$AdminAccount
    )

    if ($XMLFile) {
        ### Importing data in the LogsToAnalyze (Scope: Script) variable
        Write-Log -function "MailboxMigration - Collect-MigrationLogs" -step "Importing data from `"$XMLFile`" file, in the LogsToAnalyze variable" -Description "Success"
        $TheMigrationLogs = Import-Clixml $XMLFile
        foreach ($Log in $TheMigrationLogs) {
            $LogEntry = New-Object PSObject
            $LogEntry | Add-Member -NotePropertyName GUID -NotePropertyValue $($Log.MailboxIdentity.ObjectGuid.ToString())
            $LogEntry | Add-Member -NotePropertyName Name -NotePropertyValue $($Log.MailboxIdentity.Name.ToString())
            $LogEntry | Add-Member -NotePropertyName DistinguishedName -NotePropertyValue $($Log.MailboxIdentity.DistinguishedName.ToString())
            $LogEntry | Add-Member -NotePropertyName SID -NotePropertyValue $($Log.MailboxIdentity.SecurityIdentifierString.ToString())
            $LogEntry | Add-Member -NotePropertyName Logs -NotePropertyValue $Log

            $null = $script:LogsToAnalyze.Add($LogEntry)
        }
        $TheEnvironment = "FromFile"
        $LogFrom = "FromFile"
        $LogType = "FromFile"
        $TheMigrationType = "FromFile"
    }
    elseif ($ConnectToExchangeOnline) {
        ### Connecting to Exchange Online in order to collect the needed/correct mailbox migration logs
        #Write-Host "This part is not yet implemented" -ForegroundColor Red

        if ($MigrationType -eq "Hybrid") {
            Collect-MoveRequestStatistics -AffectedUsers $AffectedUsers
            $LogType = "MoveRequestStatistics"
        }
        elseif ($MigrationType -eq "IMAP") {
            Collect-SyncRequestStatistics -AffectedUsers $AffectedUsers
            $LogType = "SyncRequestStatistics"
        }
        elseif (($MigrationType -eq "Cutover") -or ($MigrationType -eq "Staged")) {
            Collect-MigrationUserStatistics -AffectedUsers $AffectedUsers
            $LogType = "MigrationUserStatistics"
        }
        $TheEnvironment = "Exchange Online"
        $LogFrom = "FromExchangeOnline"
        $TheMigrationType = $MigrationType
    }

    if ($script:LogsToAnalyze) {
        foreach ($LogEntry in $script:LogsToAnalyze) {
            $TheInfo = Create-MoveObject -MigrationLogs $LogEntry -TheEnvironment $TheEnvironment -LogFrom $LogFrom -LogType $LogType -MigrationType $TheMigrationType
            $null = $script:ParsedLogs.Add($TheInfo)
        }
    }
}


function Create-MoveObject {
    param (
        $MigrationLogs,

        [ValidateSet("Exchange Online", "Exchange OnPremises", "FromFile")]
        [string]$TheEnvironment,

        [ValidateSet("FromFile", "FromExchangeOnline", "FromExchangeOnPremises")]
        [string]$LogFrom,

        [ValidateSet("MoveRequestStatistics", "MoveRequest", "MigrationUserStatistics", "MigrationUser", "MigrationBatch", "SyncRequestStatistics", "SyncRequest", "MailboxStatistics", "FromFile")]
        [string]$LogType,

        [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged", "FromFile")]
        [string]$MigrationType
    )

    # List of fields to output
    [Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

    # Create the Result object that will be used to store all results
    $MoveAnalysis = New-Object PSObject
    $OrderedFields | foreach {$MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty}

    # Pull everything that we need, that is common to all status types
    $MoveAnalysis.BasicInformation        = New-BasicInformation -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.PerformanceStatistics   = New-PerformanceStatistics -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.FailureSummary          = New-FailureSummary -RequestStats $($MigrationLogs.Logs)
<#    
    ##$MoveAnalysis.FailureStatistics       = New-FailureStatistics -FailureSummaries $MoveAnalysis.FailureSummary
    ##$MoveAnalysis.LargeItemSummary        = New-LargeItemSummary -RequestStats $($MigrationLogs.Logs)
    ##$MoveAnalysis.BadItemSummary          = New-BadItemSummary -RequestStats $($MigrationLogs.Logs)
    # Add fields that are not printed in the analysis
    $MoveAnalysis | Add-Member -NotePropertyName Report -NotePropertyValue $($MigrationLogs.Logs.Report)
    $MoveAnalysis | Add-Member -NotePropertyName DiagnosticInfo -NotePropertyValue $($MigrationLogs.Logs.DiagnosticInfo)
    $timelineMonth = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Month
    $timelineDay = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Day
    $timelineHour = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Hour
    $timelineMinute = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Minute
    $Timeline = New-Object PSObject
    $Timeline | Add-Member -NotePropertyName timelineMonth -NotePropertyValue $timelineMonth
    $Timeline | Add-Member -NotePropertyName timelineDay -NotePropertyValue $timelineDay
    $Timeline | Add-Member -NotePropertyName timelineHour -NotePropertyValue $timelineHour
    $Timeline | Add-Member -NotePropertyName timelineMinute -NotePropertyValue $timelineMinute
    $MoveAnalysis | Add-Member -NotePropertyName Timeline -NotePropertyValue $Timeline
#>
    $DetailsAboutTheMove = New-Object PSObject
    $DetailsAboutTheMove | Add-Member -NotePropertyName Environment -NotePropertyValue $TheEnvironment
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogFrom -NotePropertyValue $LogFrom
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogType -NotePropertyValue $LogType
    $DetailsAboutTheMove | Add-Member -NotePropertyName MigrationType -NotePropertyValue $MigrationType
    $DetailsAboutTheMove | Add-Member -NotePropertyName PrimarySMTPAddress -NotePropertyValue $($MigrationLogs.Name)

    $MoveAnalysis | Add-Member -NotePropertyName DetailsAboutTheMove -NotePropertyValue $DetailsAboutTheMove
    return $MoveAnalysis

}


# Create the Basic Information object and populate it with the baseline values
Function New-BasicInformation
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Build all properties to be added to the oubject
    New-Object PSObject -Property ([ordered]@{
        Alias           = $RequestStats.Alias 
        ExchangeGuid    = $RequestStats.ExchangeGuid
        Created         = $RequestStats.QueuedTimestamp
        Completed       = $RequestStats.CompletionTimeStamp
        Failed          = $RequestStats.FailureTimeStamp
        Direction       = ([String]$RequestStats.Direction)
        Status          = ([String]$RequestStats.Status)
        StatusDetail    = ([String]$RequestStats.StatusDetail)
        Workload        = ([String]$RequestStats.Workloadtype)
        Flags           = ([String]$RequestStats.Flags)
        SourceServer    = $RequestStats.SourceServer
        SourceVersion   = $RequestStats.SourceVersion
        SourceDatabase  = $RequestStats.SourceDatabase
        TargetServer    = $RequestStats.TargetServer
        TargetVersion   = $RequestStats.TargetVersion
        TargetDatabase  = $RequestStats.TargetDatabase
        MRSProxy        = $RequestStats.RemoteHostName
        RemoteDatabase  = $RequestStats.RemoteDatabaseName
        Protected       = $RequestStats.Protect
        BadItemLimit    = ([int][String]$RequestStats.BadItemLimit)
        LargeItemLimit  = ([int][String]$RequestStats.LargeItemLimit)
        #Failures        = $RequestStats.Report.Failures
        BadItems        = $RequestStats.Report.BadItems
        LargeItems      = $RequestStats.Report.LargeItems
    })
}



# Build the stats for performance
Function New-PerformanceStatistics
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    New-Object PSObject -Property ([ordered]@{
        MigrationDuration         = $RequestStats.TotalInProgressDuration
        ##TotalGBTransferred        = (Get-Bytes $RequestStats.BytesTransferred) / 1GB
        PercentComplete           = $RequestStats.PercentComplete
        ##DataTransferRateMBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1MB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        ##DataTransferRateGBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1GB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        AverageSourceLatency      = Eval-Safe { $RequestStats.report.sessionstatistics.sourcelatencyinfo.average }
        AverageDestinationLatency = Eval-Safe { $RequestStats.report.sessionstatistics.destinationlatencyinfo.average }
        SourceSideDuration        = Eval-Safe { $RequestStats.Report.SessionStatistics.SourceProviderInfo.TotalDuration }
        DestinationSideDuration   = Eval-Safe { $RequestStats.Report.SessionStatistics.DestinationProviderInfo.TotalDuration }
        PercentDurationIdle       = Eval-Safe { ((DurationToSeconds $RequestStats.TotalIdleDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationSuspended  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalSuspendedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationFailed     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalFailedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationQueued     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalQueuedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationLocked     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalStalledDueToMailboxLockedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationTransient  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalTransientFailureDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
    })
}



# Creates an object with just the failure message and timestamp
Function New-FailureSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Create the object
    $compactFailures = @()

    # If we have no failures make sure we write something
    if ($RequestStats.report.failures -eq $null)
    {
        $compactFailures += New-Object PSObject -Property @{
            TimeStamp = [DateTime]::MinValue
            FailureType = "No Failures Found"
        }
    }
    # Pull out just what we want in the compact report
    else
    {
        $compactFailures += $RequestStats.report.failures | Select-object -Property TimeStamp,Failuretype,Message

        # Pull in the entries that indicate us starting a mailbox move
        $compactFailures += ($RequestStats.report.entries | where { $_.message -like "*examining the request*" } | 
        select-Object @{
            Name = "TimeStamp"; 
            Expression = { $_.CreationTime }
        },
        @{
            Name = "FailureType";
            Expression = { "-->MRSPickingUpMove" }
        },
        Message)
    }

    $compactFailures = $compactFailures | sort-Object -Property timestamp

    Return $compactFailures
}


# Evaluates an expression and returns the result
# if an exception is thrown, returns a default value
Function Eval-Safe
{
    param(
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Expression,

        [Parameter(Mandatory=$false)]
        $DefaultValue = $null
    )

    try
    {
        return (Invoke-Command -ScriptBlock $Expression)
    }
    catch
    {
        Write-Warning ("Eval-Safe: Error: '{0}'; returning default value: {1}" -f $_,$DefaultValue)
        return $DefaultValue
    }
}


# Returns totalseconds or 0 for durations
# This operates against: System.TimeSpan, M.E.Data.EnhancedTimeSpan and Deserialized.M.E.Data.EnhancedTimeSpan
Function DurationtoSeconds
{
    Param(
        [Parameter(Mandatory = $false)]
        $time = $null
    )

    if ($time -eq $null) { 0 }
    else { $time.TotalSeconds }
}


### <summary>
### Selected-ConnectToExchangeOnlineOption function is used to connect to Exchange Online, and collect from there the mailbox migration logs,
### for the affected user, by running the correct commands, based on the migration type
### </summary>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the mailbox migration logs </param>
### <param name="MigrationType">MigrationType represents the migration type for which we collect the mailbox migration logs </param>
### <param name="TheAdminAccount">TheAdminAccount represents username of an Admin that we will use in order to connect to Exchange Online </param>
function Selected-ConnectToExchangeOnlineOption {

    Connect-O365PS "EXO"

    Write-Log -function "MailboxMigration - Selected-ConnectToExchangeOnlineOption" -step "Trying to collect the AffectedUser..." -Description "Success"
    [string]$AffectedUsers = Ask-ForDetailsAboutUser -NumberOfChecks 1

    [System.Collections.ArrayList]$PrimarySMTPAddresses = @()
    $TheRecipients = Find-TheRecipient -TheEnvironment 'Exchange Online' -TheAffectedUsers $AffectedUsers

    foreach ($Recipient in $TheRecipients) {
        $null = $PrimarySMTPAddresses.Add($($Recipient.PrimarySMTPAddress))
    }

    [string]$TheAddresses = ""
    [int]$Counter = 0
    if ($($PrimarySMTPAddresses.Count) -eq 0) {
        Write-Log -function "MailboxMigration - Selected-ConnectToExchangeOnlineOption" -step "Get list of PrimarySMTPAddresses of the affected users" -Description "We were unable to find any valid SMTP Address to be used for further investigation"
        throw "We were unable to find any valid SMTP Address to be used for further investigation"
    }
    elseif ($($PrimarySMTPAddresses.Count) -eq 1) {
        $TheAddresses = $PrimarySMTPAddresses[0]
    }
    elseif ($($PrimarySMTPAddresses.Count) -gt 1) {
        foreach ($PrimarySMTPAddress in $PrimarySMTPAddresses) {
            if ($Counter -eq 0) {
                [string]$TheAddresses = $PrimarySMTPAddress
                $Counter++
            }
            elseif (($Counter -le $($PrimarySMTPAddresses.Count))) {
                [string]$TheAddresses = $TheAddresses + ", $PrimarySMTPAddress"
                $Counter++
            }
        }
    }

    Collect-MigrationLogs -ConnectToExchangeOnline -MigrationType "Hybrid" -AdminAccount $TheAdminAccount -AffectedUsers $PrimarySMTPAddresses

}


### <summary>
### Ask-ForDetailsAboutUser function is used to collect the Affected user.
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in order to provide different messages when collecting the affected user
### for the first time, or if you are re-asking for the affected user </param>
function Ask-ForDetailsAboutUser {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $NumberOfChecks
    )    

    Write-Host
    if ($NumberOfChecks -eq "1") {
        ### Asking for the affected user, for the first time
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "Asking to provide the affected user. Iteration 1" -Description "Success"
        Write-Host "Please provide the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        $NumberOfChecks++
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "The affected user provided is: $TheUserName" -Description "Success"
    }
    else {
        ### Re-asking for the affected user
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "Re-asking to provide the affected user. Iteration $NumberOfChecks" -Description "Success"
        Write-Host "Please provide again the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "The affected user provided is: $TheUserName" -Description "Success"
    }

    ### Validating if the user provided is the affected user
    Write-Host
    Write-Host "You entered " -NoNewline -ForegroundColor Cyan
    Write-Host "$TheUserName" -NoNewline -ForegroundColor White
    Write-Host " as being the affected user. Is this correct?" -ForegroundColor Cyan
    Write-Host "`t[Y] Yes     [N] No      (default is `"Y`"): " -NoNewline -ForegroundColor White
    $ReadFromKeyboard = Read-Host

    [bool]$TheKey = $true
    Switch ($ReadFromKeyboard) 
    { 
      Y {$TheKey=$true} 
      N {$TheKey=$false} 
      Default {$TheKey=$true} 
    }

    if ($TheKey) {
        ### Received confirmation that the user provided is the affected user.
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "Got confirmation that `"$TheUserName`" is indeed the affected user" -Description "Success"
    }
    else {
        ### The user provided is not the affected user. Asking again for the affected user.
        Write-Log -function "MailboxMigration - Ask-ForDetailsAboutUser" -step "`"$TheUserName`" is not the affected user. Starting over the process of asking for the affected user" -Description "Success"
        [string]$TheUserName = Ask-ForDetailsAboutUser -NumberOfChecks $NumberOfChecks
    }

    ### The function will return the affected user
    return $TheUserName
}



function Find-TheRecipient {
    [CmdletBinding()]
    Param (
        [ValidateSet("Exchange Online", "Exchange OnPremises")]
        [string]
        $TheEnvironment,
        [string[]]
        $TheAffectedUsers
    )

    [System.Collections.ArrayList]$Recipients = @()
    foreach ($User in $TheAffectedUsers) {
        $TheCommand = Create-CommandToInvoke -TheEnvironment $TheEnvironment -CommandFor "Recipient"
        try {
            Write-Log -function "MailboxMigration - Find-TheRecipient" -step "Collecting `"Get-Recipient`" for `"$User`"" -Description "Success"
            $ExpressionResults = Invoke-Expression $($TheCommand.FullCommand)
            Write-Log -function "MailboxMigration - Find-TheRecipient" -step "We were able to identify the recipient in $TheEnvironment for `"$User`".`n`tPrimarySmtpAddress:`t$($ExpressionResults.PrimarySmtpAddress)`n`tExchangeGuid:`t`t$($ExpressionResults.ExchangeGuid)`n`tRecipientType:`t`t$($ExpressionResults.RecipientType)`n`tRecipientTypeDetails:`t$($ExpressionResults.RecipientTypeDetails)" -Description "Success"
            Write-Log -function "MailboxMigration - Find-TheRecipient" -step "From now on, we will use its PrimarySMTPAddress, `"$($ExpressionResults.PrimarySmtpAddress)`", when providing details about `"$User`"" -Description "Success"

            $null = $Recipients.Add($ExpressionResults)
        }
        catch {
            Write-Log -function "MailboxMigration - Find-TheRecipient" -step "Unable to identify the Recipient using information you provided (`"$User`")" -Description "Success"
        }
    }

    if ($($Recipients.Count) -eq 0){
        Write-Log -function "MailboxMigration - Find-TheRecipient" -step "No recipients in the Organization" -Description "We were unable to identify any Recipients in your organization, for the users you provided"
        throw "We were unable to identify any Recipients in your organization, for the users you provided"
    }
    else {
        return $Recipients
    }

}



### <summary>
### Create-CommandToInvoke function is used to create the exact command to run, in order to collect the correct migration logs
### </summary>
### <param name="TheEnvironment">TheEnvironment represents the environment in which the command will run </param>
function Create-CommandToInvoke {
    param (
        [ValidateSet("Exchange Online", "Exchange OnPremises")]
        [string]
        $TheEnvironment,
        [ValidateSet("MoveRequestStatistics", "MoveRequest", "MigrationUserStatistics", "MigrationUser", "MigrationBatch", "SyncRequestStatistics", "SyncRequest", "MailboxStatistics", "Recipient")]
        [string]
        $CommandFor
    )

    $TheResultantCommand = New-Object PSObject

    if ($TheEnvironment -eq "Exchange Online") {
        if ($CommandFor -eq "MoveRequestStatistics") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MoveRequestStatistics")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "MoveRequestStatistics `$User -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`" -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "MoveRequest") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MoveRequest")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "MoveRequest `$User -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "MigrationUserStatistics") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MigrationUserStatistics")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "MigrationUserStatistics `$User -IncludeSkippedItems -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`" -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "MigrationUser") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MigrationUser")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "MigrationUser `$User -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "MigrationBatch") {
            <#
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MigrationBatch")
            [string]$TheCommand = "(Get-"+ $script:EXOCommandsPrefix + "MigrationBatch `$User -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`" -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
            #>
        }
        elseif ($CommandFor -eq "SyncRequestStatistics") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "SyncRequestStatistics")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "SyncRequestStatistics `$User -IncludeReport -DiagnosticInfo `"showtimeslots, showtimeline, verbose`" -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "SyncRequest") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "SyncRequest")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "SyncRequest -Mailbox `$User -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "MailboxStatistics") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "MailboxStatistics")
            [string]$TheCommand = "(Get-"+ $script:EXOCommandsPrefix + "MailboxStatistics `$User -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop).MoveHistory | where {[string]`$(`$_.WorkloadType) -eq `"Onboarding`"} | select -First 1"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "Recipient") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:EXOCommandsPrefix + "Recipient")
            [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "Recipient `$User -ResultSize Unlimited -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
    }
    else {
        if ($CommandFor -eq "MailboxStatistics") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:ExOnPremCommandsPrefix + "MailboxStatistics")
            [string]$TheCommand = "(Get-"+ $script:ExOnPremCommandsPrefix + "MailboxStatistics `$User -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop).MoveHistory | where {[string]`$(`$_.WorkloadType) -eq `"Offboarding`"} | select -First 1"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand
        }
        elseif ($CommandFor -eq "Recipient") {
            $TheResultantCommand | Add-Member -NotePropertyName Command -NotePropertyValue ("Get-"+ $script:ExOnPremCommandsPrefix + "Recipient")
            [string]$TheCommand = "Get-"+ $script:ExOnPremCommandsPrefix + "Recipient `$User -ResultSize Unlimited -ErrorAction Stop"
            $TheResultantCommand | Add-Member -NotePropertyName FullCommand -NotePropertyValue $TheCommand            
        }
    }

    return $TheResultantCommand
}



function Collect-MoveRequestStatistics {
    param (
        [string[]]
        $AffectedUsers
    )

    Write-Log -function "MailboxMigration - Collect-MoveRequestStatistics" -step "Collecting Get-MoveRequestStatistics for each Affected users" -Description "Success"
    $TheCommand = Create-CommandToInvoke -TheEnvironment 'Exchange Online' -CommandFor "MoveRequestStatistics"

    if ($($TheCommand.Command)) {
        foreach ($User in $AffectedUsers) {
            try {
                $null = Get-Command $($TheCommand.Command) -ErrorAction Stop
                Write-Log -function "MailboxMigration - Collect-MoveRequestStatistics" -step "Running the following command:`n`t$($TheCommand.FullCommand.Replace("`$User", "$User"))" -Description "Success"
                try {
                    $ExpressionResults = Invoke-Expression $($TheCommand.FullCommand)
                    Write-Log -function "MailboxMigration - Collect-MoveRequestStatistics" -step "MoveRequestStatistics successfully collected for `"$User`" user" -Description "Success"
                    $LogEntry = New-Object PSObject
                    $LogEntry | Add-Member -NotePropertyName PrimarySMTPAddress -NotePropertyValue $User
                    $LogEntry | Add-Member -NotePropertyName MigrationType -NotePropertyValue "Hybrid"
                    $LogEntry | Add-Member -NotePropertyName LogType -NotePropertyValue "MoveRequestStatistics"
                    $LogEntry | Add-Member -NotePropertyName Logs -NotePropertyValue $ExpressionResults
                    $void = $script:LogsToAnalyze.Add($LogEntry)
                    [string]$Path = $script:TheWorkingDirectorySavedData + "\EXO_MoveRequestStatistics_" + [string]$User + ".xml"
                    $LogEntry | Export-Clixml $Path -Force
                }
                catch {
                    Write-Log -function "MailboxMigration - Collect-MoveRequestStatistics" -step "We were unable to collect MoveRequestStatistics for `"$User`" user" -Description "Error"
                }
            }
            catch {                
                Write-Log -function "MailboxMigration - Collect-MoveRequestStatistics" -step "You do not have permissions to run `"$($TheCommand.Command)`" command" -Description "Error"
                #Collect-MoveRequest -AffectedUsers $AffectedUsers
            }
        }
    }
}


### <summary>
### Export-MailboxMigrationReportToHTML function is used to create the object that will be converted to HTML report
### </summary>
function Export-MailboxMigrationReportToHTML {

    [System.Collections.ArrayList]$TheObjectToConvertToHTML = @()

    foreach ($Entry in $script:ParsedLogs) {

        [string]$SectionTitle = "Basic Information"
        [string]$Description = "Below are the `"Basic Information`" for $($Entry.BasicInformation.Alias)'s migration`<p`>`&nbsp*`&nbspThe title of this section is colored in Red in case the <u>Status</u> of the migration is <u>Failed</u>"
        [PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $($Entry.BasicInformation) -TableType "List"
        $null = $TheObjectToConvertToHTML.Add($TheCommand)

        [string]$SectionTitle = "Performance Statistics"
        [string]$Description = "Below are the `"Performance Statistics`" for $($Entry.BasicInformation.Alias)'s migration"
        [PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor "Black" -Description $Description -DataType "ArrayList" -EffectiveDataArrayList $($Entry.PerformanceStatistics) -TableType "List"
        $null = $TheObjectToConvertToHTML.Add($TheCommand)

    }

    [string]$FilePath = $ExportPath + "\MailboxMigration_Hybrid_SummaryReport.html"
    Export-ReportToHTML -FilePath $FilePath -PageTitle "Mailbox Migration Report" -ReportTitle "Mailbox Migration - Hybrid - Summary Report" -TheObjectToConvertToHTML $TheObjectToConvertToHTML
}







function JustFor-Testing {
    Write-Host
    Write-Host "Details from the mailbox migration log:" -ForegroundColor Green

    foreach ($Entry in $script:ParsedLogs) {
        if ($($Entry.DetailsAboutTheMove.PrimarySMTPAddress) -eq "FromFile")
        {
            #Write-Log "[INFO] || Providing details of the migration from the file you provided"
            Write-Host "[INFO] || Providing details of the migration from the file you provided"
        }
        else {
            #Write-Log "[INFO] || Providing details of the migration for $($Entry.DetailsAboutTheMove.PrimarySMTPAddress)"
            Write-Host "[INFO] || Providing details of the migration for $($Entry.DetailsAboutTheMove.PrimarySMTPAddress)"
        }

        if ($($Entry.DetailsAboutTheMove.PrimarySMTPAddress)) {
            Write-Host "PrimarySMTPAddress: " -ForegroundColor Cyan
            Write-Host "$($Entry.DetailsAboutTheMove.PrimarySMTPAddress)" -ForegroundColor White
            #Write-Log ("PrimarySMTPAddress: `n$($Entry.DetailsAboutTheMove.PrimarySMTPAddress)") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.BasicInformation)) {
            Write-Host "Basic Information: " -ForegroundColor Cyan
            $($Entry.BasicInformation)
            #Write-Log ("Basic Information: `n$($Entry.BasicInformation)") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.PerformanceStatistics)) {
            Write-Host "Performance Statistics: " -ForegroundColor Cyan
            $($Entry.PerformanceStatistics)
            #Write-Log ("Performance Statistics: `n$($Entry.PerformanceStatistics)") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineMinute)) {
            Write-Host "Details about timeline, by minute: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineMinute)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            #Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineHour)) {
            Write-Host "Details about timeline, by hour: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineHour)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            #Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineDay)) {
            Write-Host "Details about timeline, by Day: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineDay)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            <#foreach ($timelineMonthSortedEntry in $TheEntriesToDisplay) {
                Write-Host 
                Write-Host "`t$($timelineMonthSortedEntry.State): " -ForegroundColor Cyan -NoNewline
                Write-Host "`t$($timelineMonthSortedEntry.Milliseconds)" -ForegroundColor White -NoNewline
                $ThePercent = (([int]$($timelineMonthSortedEntry.Milliseconds)/[int]$TheDurationMilliseconds)*100).ToString("#.##")
                Write-Host " ($ThePercent `%)"
            }#>
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            #Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineMonth)) {
            Write-Host "Details about timeline, by month: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineMonth)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            <#foreach ($timelineMonthSortedEntry in $TheEntriesToDisplay) {
                Write-Host 
                Write-Host "`t$($timelineMonthSortedEntry.State): " -ForegroundColor Cyan -NoNewline
                Write-Host "`t$($timelineMonthSortedEntry.Milliseconds)" -ForegroundColor White -NoNewline
                $ThePercent = (([int]$($timelineMonthSortedEntry.Milliseconds)/[int]$TheDurationMilliseconds)*100).ToString("#.##")
                Write-Host " ($ThePercent `%)"
            }#>
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            #Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        Write-Host
    }
}

function Start-MailboxMigrationMainScript {

    Write-Log -function "MailboxMigration - Start-MailboxMigrationMainScript" -step "Show-MailboxMigrationMenu" -Description "Success"
    Show-MailboxMigrationMenu

    Write-Log -function "MailboxMigration - Start-MailboxMigrationMainScript" -step "Export-MailboxMigrationReportToHTML" -Description "Success"
    Export-MailboxMigrationReportToHTML

    Write-Log -function "MailboxMigration - Start-MailboxMigrationMainScript" -step "Read-Key" -Description "Success"
    Read-Key

    Write-Log -function "MailboxMigration - Start-MailboxMigrationMainScript" -step "Start-O365TroubleshootersMenu" -Description "Success"
    Start-O365TroubleshootersMenu

}

#endregion Functions


###############
# Main script #
###############
#region Main script

try {
    Write-Log -function "MailboxMigration" -step "Start-MailboxMigrationMainScript" -Description "Success"
    Start-MailboxMigrationMainScript
}
catch {
    Write-Log -function "MailboxMigration" -step "MainScript" -Description "$_"
    Write-Log -function "MailboxMigration" -step "MainScript" -Description "Error. Script will now exit"
    Write-Host "[ERROR] || $_" -ForegroundColor Red
    Write-Host "[ERROR] || Script will now exit" -ForegroundColor Red
}

#endregion Main script