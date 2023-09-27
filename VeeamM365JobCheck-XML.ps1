<#
    .SYNOPSIS
        This script checks the status of all active Veeam for Microsoft 365 Jobs on a backup server.
        It collects detailed information and creates an XML file per backupjob as output.
        The XML will be placeed in C:\Temp\VeeamResults where it can be retreived by the PRTG-Sensor

    .INPUTS
        None

    .OUTPUTS
        The script creates a XML file formated for PRTG.

    .LINK
        Disclamer: https://raw.githubusercontent.com/tn-ict/Public/master/Disclaimer/DISCLAIMER

    .NOTES
        Author:  Andreas Bucher
        Version: 1.0.0
        Date:    27.09.2023
        Purpose: XML-Part of the PRTG-Sensor VeeamM365JobCheck

    .EXAMPLE
        powershell.exe -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File "C:\Script\VeeamM365JobCheck-XML.ps1"
        
        Run this script with task scheduler use powershell.exe as program and the parameters as described
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Use TLS1.2 for Invoke-Webrequest
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# General parameters
$nl               = [Environment]::NewLine
$resultFolder     = "C:\Temp\VeeamResults"

# PRTG parameters
$WarningLevel = 24 # Warninglevel in hours for last backup session
$ErrorLevel   = 36 # Errorlevel in hours for last backup session

# Define JobResult object and parameters
$JobResult = [PSCustomObject]@{
    Name      = ""
    Value     = 0
    Text      = ""
    Warning   = 0
    Error     = 0
    CountObj  = 0
    ProcItems = 0
    ProcRate  = 0
    ReadRate  = 0
    WriteRate = 0
    TransData = 0
    Duration  = 0
    LastBkp   = 0
}

#-----------------------------------------------------------[Functions]------------------------------------------------------------
# Export XML
function Set-XMLContent {
    param(
        $JobResult
    )

    # Create XML-Content
    $result= ""
    $result+= '<?xml version="1.0" encoding="UTF-8" ?>' + $nl
    $result+= "<prtg>" + $nl

    $result+=   "<Error>$($JobResult.Error)</Error>" + $nl
    $result+=   "<Text>$($JobResult.Text)</Text>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Status</channel>" + $nl
    $result+=   "  <value>$($JobResult.Value)</value>" + $nl
    $result+=   "  <Warning>$($JobResult.Warning)</Warning>" + $nl
    $result+=   "  <LimitMaxWarning>2</LimitMaxWarning>" + $nl
    $result+=   "  <LimitMaxError>3</LimitMaxError>" + $nl
    $result+=   "  <LimitMode>1</LimitMode>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Abgearbeitete Objekte</channel>" + $nl
    $result+=   "  <value>$($JobResult.CountObj)</value>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Items/s</channel>" + $nl
    $result+=   "  <value>$($JobResult.ProcItems)</value>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Processing Rate</channel>" + $nl
    $result+=   "  <value>$($JobResult.ProcRate)</value>" + $nl
    $result+=   "  <Float>1</Float>" + $nl
    $result+=   "  <DecimalMode>Auto</DecimalMode>" + $nl
    $result+=   "  <CustomUnit>MB/s</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Read Rate</channel>" + $nl
    $result+=   "  <value>$($JobResult.ReadRate)</value>" + $nl
    $result+=   "  <Float>1</Float>" + $nl
    $result+=   "  <DecimalMode>Auto</DecimalMode>" + $nl
    $result+=   "  <CustomUnit>MB/s</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Write Rate</channel>" + $nl
    $result+=   "  <value>$($JobResult.WriteRate)</value>" + $nl
    $result+=   "  <Float>1</Float>" + $nl
    $result+=   "  <DecimalMode>Auto</DecimalMode>" + $nl
    $result+=   "  <CustomUnit>MB/s</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Transferiert</channel>" + $nl
    $result+=   "  <value>$($JobResult.TransData)</value>" + $nl
    $result+=   "  <Float>1</Float>" + $nl
    $result+=   "  <DecimalMode>Auto</DecimalMode>" + $nl
    $result+=   "  <CustomUnit>MB</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Dauer</channel>" + $nl
    $result+=   "  <value>$($JobResult.Duration)</value>" + $nl
    $result+=   "  <Float>1</Float>" + $nl
    $result+=   "  <DecimalMode>Auto</DecimalMode>" + $nl
    $result+=   "  <CustomUnit>Min</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "</result>" + $nl

    $result+=   "<result>" + $nl
    $result+=   "  <channel>Stunden seit letzem Job</channel>" + $nl
    $result+=   "  <value>$($JobResult.LastBkp)</value>" + $nl
    $result+=   "  <CustomUnit>h</CustomUnit>" + $nl
    $result+=   "  <showChart>1</showChart>" + $nl
    $result+=   "  <showTable>1</showTable>" + $nl
    $result+=   "  <LimitMaxWarning>$WarningLevel</LimitMaxWarning>" + $nl
    $result+=   "  <LimitWarningMsg>Backup-Job älter als 24h</LimitWarningMsg>" + $nl
    $result+=   "  <LimitMaxError>$ErrorLevel</LimitMaxError>" + $nl
    $result+=   "  <LimitErrorMsg>Backup-Job älter als 36h</LimitErrorMsg>" + $nl
    $result+=   "  <LimitMode>1</LimitMode>" + $nl
    $result+=   "</result>" + $nl

    $result+= "</prtg>" + $nl

    # Write XML-File
    if(-not (test-path $resultFolder)){ New-Item -Path $resultFolder -ItemType Directory }
    $xmlFilePath = "$resultFolder\$($JobResult.Name).xml"
    $result | Out-File $xmlFilePath -Encoding utf8

}
# Calculate Backup job details
function Get-JobResult {
    param(
        $Session
    )

    # Get session processed objects and duration
    $JobResult.CountObj = $Session.Statistics.ProcessedObjects
    $JobResult.Duration = [Math]::Round(($Session.EndTime - $Session.CreationTime).TotalMinutes, 2)

    # Get session status
    if     ($Session.Status -eq "Success")        { $JobResult.Value = 1; $JobResult.Warning = 0; $JobResult.Error = 0; $JobResult.Text = "BackupJob $($JobResult.Name) erfolgreich" }
    elseif ($Session.Status -eq "Warning")        { $JobResult.Value = 2; $JobResult.Warning = 1; $JobResult.Error = 0; $JobResult.Text = "BackupJob $($JobResult.Name) Warnung. Bitte pr&#252;fen" }
    elseif ($Session.Status -eq "Failed")         { $JobResult.Value = 3; $JobResult.Warning = 0; $JobResult.Error = 1; $JobResult.Text = "BackupJob $($JobResult.Name) fehlerhaft" }
    elseif ($Session.Status -eq "Running")        { $JobResult.Value = 2; $JobResult.Warning = 1; $JobResult.Error = 0; $JobResult.Text = "BackupJob $($JobResult.Name) l&#228;uft noch" }
    elseif ($Session.Status -eq "Disconnected")   { $JobResult.Value = 3; $JobResult.Warning = 0; $JobResult.Error = 1; $JobResult.Text = "BackupJob $($JobResult.Name) disconnected" }
    elseif ($Session.Status -eq "Stopped")        { $JobResult.Value = 1; $JobResult.Warning = 0; $JobResult.Error = 0; $JobResult.Text = "BackupJob $($JobResult.Name) gestoppt" }
    elseif ($Session.Status -eq "Queued")         { $JobResult.Value = 1; $JobResult.Warning = 0; $JobResult.Error = 0; $JobResult.Text = "BackupJob $($JobResult.Name) eingereiht" }
    elseif ($Session.Status -eq "NotConfigured")  { $JobResult.Value = 3; $JobResult.Warning = 0; $JobResult.Error = 1; $JobResult.Text = "BackupJob $($JobResult.Name) nicht konfiguriert" }
    else                                          { $JobResult.Value = 3; $JobResult.Warning = 0; $JobResult.Error = 1; $JobResult.Text = "BackupJob $($JobResult.Name) unbekannter Fehler" }

    Return $JobResult
}
# Get Job Statistics
function Get-Statistics {
    param(
        $Statistic
    )

    # Get item per second
    $procrate = $Session.Statistics.ProcessingRate -split '\s+'
    $JobResult.ProcItems = $procrate[2] -replace '[()]',''

    # Check processing speed and output it as MB/s
    $procrate = $Session.Statistics.ProcessingRate -split '\s+'
    if ($procrate[1] -eq "B/s")  { $JobResult.ProcRate = [Math]::Round([Decimal]$procrate[0]/1MB, 2) }
    if ($procrate[1] -eq "KB/s") { $JobResult.ProcRate = [Math]::Round([Decimal]$procrate[0]/1KB, 2) }
    if ($procrate[1] -eq "MB/s") { $JobResult.ProcRate = [Math]::Round([Decimal]$procrate[0], 2) }

    # Check read speed and output it as MB/s
    $readrate = $Session.Statistics.ReadRate -split '\s+'
    if ($readrate[1] -eq "B/s")  { $JobResult.ReadRate = [Math]::Round([Decimal]$readrate[0]/1MB, 2) }
    if ($readrate[1] -eq "KB/s") { $JobResult.ReadRate = [Math]::Round([Decimal]$readrate[0]/1KB, 2) }
    if ($readrate[1] -eq "MB/s") { $JobResult.ReadRate = [Math]::Round([Decimal]$readrate[0], 2) }

    # Check write speed and output it as MB/s
    $writerate = $Session.Statistics.WriteRate -split '\s+'
    if ($writerate[1] -eq "B/s")  { $JobResult.WriteRate = [Math]::Round([Decimal]$writerate[0]/1MB, 2) }
    if ($writerate[1] -eq "KB/s") { $JobResult.WriteRate = [Math]::Round([Decimal]$writerate[0]/1KB, 2) }
    if ($writerate[1] -eq "MB/s") { $JobResult.WriteRate = [Math]::Round([Decimal]$writerate[0], 2) }

    # # Check transfered speed and output it as MB/s
    $transfered = $Session.Statistics.TransferredData -split '\s+'
    if ($transfered[1] -eq "B")  { $JobResult.TransData = [Math]::Round([Decimal]$transfered[0]/1MB, 2) }
    if ($transfered[1] -eq "KB") { $JobResult.TransData = [Math]::Round([Decimal]$transfered[0]/1KB, 2) }
    if ($transfered[1] -eq "MB") { $JobResult.TransData = [Math]::Round([Decimal]$transfered[0], 2) }
    if ($transfered[1] -eq "GB") { $JobResult.TransData = [Math]::Round([Decimal]$transfered[0]*1KB, 2) }

    Return $JobResult
}
# Get Logs for Tape, Agent and Copy-Jobs
function Get-JobLog {
    param(
        $Session
    )

    # Find warning and error messages in session log
    $warningmsg = ""
    $warningmsg = $Session.Log | Where-Object {$_.Title -like "*Warning*"} | ForEach-Object { $_.title }
    $failedmsg  = ""
    $failedmsg  = $Session.Log | Where-Object {$_.Title -like "*Failed*"} | ForEach-Object { $_.title }

    if ($failedmsg)      { Return $failedmsg }
    elseif ($warningmsg) { Return $warningmsg }
    else                 { Return }
}
#-----------------------------------------------------------[Execute]------------------------------------------------------------
# Get M365 BackupJobs
$BackupJobs = Get-VBOJob | where-object { $_.IsEnabled -eq $True }

#### Get BackupJob details ######################################################################################################
foreach($item in $BackupJobs) {

    $JobResult.Name = $item.Name

    # Load last session
    $Session = Get-VBOJobSession -Job $item -Last

    # Check job results
    $JobResult = Get-JobResult $Session
    $JobResult = Get-Statistics $Session
    $JobResult.LastBkp = (New-TimeSpan -Start $Session.CreationTime -End (Get-Date)).Hours

    # Check for messages in session log
    $CheckJobError = Get-JobLog $Session
    if ($CheckJobError) { $JobResult.Text = $CheckJobError }

    # Create XML 
    Set-XMLContent -JobResult $JobResult
}
