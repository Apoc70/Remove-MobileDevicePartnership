<#
    .SYNOPSIS
    Remove Exchange Server 2013+ Mobile Device Partnerships 
   
    Thomas Stensitzki / Sebastian Rubertus
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Send comments and remarks to: support@granikos.eu
	
    Version 1.1, 2017-02-22
     
    .LINK  
    https://www.granikos.eu/en/justcantgetenough/PostId/262/cleanup-mobile-device-partnerships
  	
    .DESCRIPTION

    This script removes mobile device association from user mailboxes that have been inactive for more than X days.

    Use the settings.xml to configure your email server settings and the min number of days for old mobile devices.

    .NOTES 
    Requirements 
    - Exchange Server 2013, 2016
    - Windows Server 2012 R2  

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release
    1.1     ReportOnly switch added (https://github.com/Apoc70/Remove-MobileDevicePartnership/issues/1) 

    This script is the successor of the ActiveSyncDevicePartnership.ps1 script which is intended to work with Exchange Server 2010
    
    .PARAMETER SendMail
    Send the list of found mobile devices by email. Email settings are controlled by a dedicated settings.xml file. See script link for more details.

    .PARAMETER ReportOnly
    Just create a report for all found mobile devices, but DO NOT DELETE the mobile device partnerships.

    .EXAMPLE
    Remove old mobile device partnerships without sending a report email

    .\Remove-MobileDevicePartnership.ps1 

    .EXAMPLE
    Remove old mobile device partnerships and send a report email

    .\Remove-MobileDevicePartnership.ps1 -SendMail

    .EXAMPLE
    Search for old mobile device partnerships and write result as CSV to disk

    .\Remove-MobileDevicePartnership.ps1 -ReportOnly
  	
#>
[CmdletBinding()]
param(
  [parameter()]
  [switch]$SendMail,
  [parameter()]
  [switch]$ReportOnly
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logfile = ('{0}\{1}_MobileDeviceCleanup.log' -f $ScriptDir, (Get-Date -format yyyy-MM-dd_HH-mm-ss))
$objCollection = @()
$timeFormat = 'yyyy-MM-dd HH:mm:ss'
$reportFilename = "MobileDevicePartnerships_$(Get-Date -f yyyyMMdd).csv"

# Import Settings.xml config file
[xml]$ConfigFile = Get-Content -Path ('{0}\Settings.xml' -f $ScriptDir)

# Email settings from Settings.xml
$smtpsettings = @{
  To = $ConfigFile.Settings.EmailSettings.MailTo
  From = $ConfigFile.Settings.EmailSettings.MailFrom
  SmtpServer = $ConfigFile.Settings.EmailSettings.SMTPServer
}

# Fetch config values from settings.xml
$LastSync = [int]$ConfigFile.Settings.OtherSettings.LastSyncDays
$MobileDeviceLimit = [int]$ConfigFile.Settings.OtherSettings.MobileDeviceLimit
$AgedDeviceLimit = [int]$ConfigFile.Settings.OtherSettings.AgedDeviceLimit

Write-Host "Removing mobile devices haven't synchronized for $($LastSync) days or more"

Function Log
{
  [CmdletBinding()]
  Param (
    [string]$logstring = ''
  )
  Add-content -Path $logfile -Value ('{0} {1} ' -f (get-date -format 'yyyy-MM-dd HH-mm-ss'), $logstring)
}

# Create a new log file
##Write-Host
#Write-Host 'Script started, creating Log File.'
Log -logstring 'Script started.'
#Write-Host

# Query User Mailboxes and Device Statistics
# Write-Host 'Querying User Mailboxes, please wait a few seconds...' -ForeGroundColor green
Log -logstring 'Querying User Mailboxes.'
# Write-Host

# Fetch all mailboxes, Get-MailboxDatabase used to query Exchange 2013+ mailboxes only.
Write-Progress -Activity "Get user mailboxes from Exchange Server" -Status "Fetching data..." -PercentComplete 10

$Mailboxes = Get-MailboxDatabase | Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -WarningAction SilentlyContinue | Sort-Object -Property DisplayName
$NumberOfMailboxes = ($Mailboxes | Measure-Object).Count
$Counter = 1
$FailedUsers = 0

Write-Host ('Number of Mailboxes: {0} ' -f $NumberOfMailboxes)
Log -logstring ('Number of Mailboxes: {0} ' -f $NumberOfMailboxes)
Write-Host

# Iterate each User Mailbox
ForEach ($Mailbox in $Mailboxes)
{
  Write-Progress -Activity "Processing mailboxes Mailbox ($Counter/$NumberOfMailboxes) | Failed: $($FailedUsers)" -Status "Mailbox: $($Mailbox.DisplayName)" -PercentComplete (($Counter/$NumberOfMailboxes)*100)

  $MailboxAlias = $Mailbox.Alias

  # Fetch all devices for a user
  try {
    $AllDevicesFromSpecificUser = Get-MobileDevice -Mailbox $MailboxAlias -Result Unlimited  -WarningAction SilentlyContinue | Get-MobileDeviceStatistics -WarningAction SilentlyContinue
    $AllOldMobileDevices = Get-MobileDevice -Mailbox $MailboxAlias -Result Unlimited  -WarningAction SilentlyContinue | Get-MobileDeviceStatistics  -WarningAction SilentlyContinue | Where-Object {$_.LastSuccessSync -le (Get-Date).AddDays($LastSync)}
  }
  catch {
    $AllDevicesFromSpecificUser = -1
    $AllOldMobileDevices = -1
    $FailedUsers++
  }

  $CountAllDevicesFromSpecificUser = $AllDevicesFromSpecificUser.Count
  $CountOldMobileDevices = $AllOldMobileDevices.Count
   
  if ($CountAllDevicesFromSpecificUser -lt $MobileDeviceLimit) {

    Write-Verbose -Message ('User {0} has only {1} mobile device(s). Nothing to delete!' -f $MailboxAlias, $CountAllDevicesFromSpecificUser)
    Log -logstring ('User {0} has only {1} mobile device(s). Nothing to delete!' -f $MailboxAlias, $CountAllDevicesFromSpecificUser)
  }
  elseif (($CountAllDevicesFromSpecificUser -ge $MobileDeviceLimit) -and ($CountOldMobileDevices -ge $AgedDeviceLimit)) {

    # If (($CountAllDevicesFromSpecificUser -gt 4) -and ($CountOldMobileDevices -gt 1))  
    Write-Host ('User {0} has {1} devices. {2} have not synced for more than {3} days.' -f $MailboxAlias, $CountAllDevicesFromSpecificUser, $CountOldMobileDevices, $LastSync) -ForegroundColor Red
    Log -logstring ('User {0} has {1} devices. {2} have not synced for more than {3} days.' -f $MailboxAlias, $CountAllDevicesFromSpecificUser, $CountOldMobileDevices, $LastSync)

    ForEach ($Device in $AllOldMobileDevices)
    {
      $ref = 0
      $DeviceType = $Device.DeviceType
      $DeviceFriendly = $Device.FriendlyName
      $DeviceID = $Device.DeviceID
      
      if([DateTime]::TryParse($Device.LastSuccessSync, [ref]$ref)) {
        $DeviceFirstSyncTime = ([DateTime]::Parse([DateTime]$Device.FirstSyncTime)).ToString($timeFormat)
      }

      if([DateTime]::TryParse($Device.LastSuccessSync, [ref]$ref)) {
        $DeviceLastSuccessSync = ([DateTime]::Parse([DateTime]$Device.LastSuccessSync)).ToString($timeFormat)
      }
      

      Write-Host
      Write-Host 'Mobile Device Properties'
      Write-Host '------------------------------------------------------'
      Write-Host ('Type         : {0} ' -f ($DeviceType))           
      Write-Host ('Friendly Name: {0} ' -f $DeviceFriendly)
      Write-Host ('ID           : {0} ' -f $DeviceID)
      Write-Host ('First Sync   : {0} ' -f $DeviceFirstSyncTime)
      Write-Host ('Last Sync    : {0} ' -f $DeviceLastSuccessSync) -ForegroundColor Red
      Log -logstring ('Removing Device {0} with ID {1} for user {2}' -f $DeviceType, $DeviceID, $MailboxAlias)
      Write-Host ('Removing Device {0}' -f $DeviceID) -ForegroundColor Red

      if(!($ReportOnly)) {  
        # DO not remove, if we want to create a report           
        # Comment following line for development purposes
        $Device | Remove-MobileDevice -WarningAction SilentlyContinue
      }

      # Add removed device to object collection for email reporting
      $obj = New-Object -TypeName PSObject
      $obj | Add-Member -MemberType NoteProperty -Name "MailboxAlias" -Value $($MailboxAlias)
      $obj | Add-Member -MemberType NoteProperty -Name "FriendlyName" -Value $($DeviceFriendly)
      $obj | Add-Member -MemberType NoteProperty -Name "Type" -Value $($DeviceType)
      $obj | Add-Member -MemberType NoteProperty -Name "ID" -Value $($DeviceID)
      $obj | Add-Member -MemberType NoteProperty -Name "FirstSyncTime" -Value $($DeviceFirstSyncTime)
      $obj | Add-Member -MemberType NoteProperty -Name "LastSyncAttemptTime" -Value $($DeviceLastSuccessSync)

      # Add object to collection
      $objCollection += $obj
    }
  }

  $Counter++

}

# do we need to write a report file to disk??
If($ReportOnly) {
  $objCollection | Export-Csv -Path (Join-Path -Path $ScriptDir -ChildPath $reportFilename) -NoTypeInformation -Encoding UTF8 -Force
}

$HtmlReport = ""

# Do we need to send an email report?
if($SendMail) {

  # Report Timestamp
  $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm-ss"

  # Some CSS to get a pretty report
  $head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($ReportTitle)</title>
<style type="text/css">
<!-
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h2{ 
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
    clear: both; 
    font-size: 100%;
    color:#354B5E; 
}
h3{
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
    border-collapse: collapse;
    border: 1px solid black;
    font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    color: black;
    margin-bottom: 10px;
}
 
table td{
    border: 1px solid black;
    font-size: 12px;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}
 
table th {
    border: 1px solid black;
    font-size: 12px;
    font-weight: bold;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}

TR:Hover TD {Background-Color: #C1D5F8;}

->
</style>
"@

  try {
    # Build message subject
    $MessageSubject = "Removed Mobile Devices Report - $($timestamp)"

    # Build Html email message
    if(($objCollection | Measure-Object).Count -ne 0) {
      [string]$HtmlReport = $objCollection | Select-Object -Property * | ConvertTo-Html -Head $head -PreContent "<h2>$($MessageSubject)</h2>" 
    }
    else {
      # Ooops, we did not find any mobile devices
      [string]$HtmlReport = 'No mobile devices found for removal.'
    }

    # try to send email
    Send-MailMessage @smtpsettings -Subject $MessageSubject -Body $HtmlReport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) # -ErrorAction STOP
  }
  catch{
    $ErrorString = $_.Exception.Message
    Write-Warning -Message $ErrorString

  }   
}

# Script finished
Write-Host
Write-Host 'Script finished!'
Write-Host
Log -logstring 'Script finished!'