# Remove-MobileDevicePartnership
Remove Exchange Server 2013+ Mobile Device Partnerships 

## Description
This script removes mobile device association from user mailboxes that have been inactive for more than X days.

Use the settings.xml to configure your email server settings and the min number of days for old mobile devices.

## Prameters
### SendMail
Send the list of found mobile devices by email. Email settings are controlled by a dedicated settings.xml file. See script link for more details.

### ReportOnly
Just create a report for all found mobile devices, but DO NOT DELETE the mobile device partnerships.

## Outputs
Sends an HTML email report if SendMail switch has been selected.

Writes a CSV file into script directory if ReportOnly switch has been selected.

## Examples
```
.\Remove-MobileDevicePartnership.ps1
```
Remove old mobile device partnerships without sending a report email

```
.\Remove-MobileDevicePartnership.ps1 -SendMail
```
Remove old mobile device partnerships and send a report email

```
.\Remove-MobileDevicePartnership.ps1 -ReportOnly
```
Search for old mobile device partnerships and write result as CSV to disk

## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Cleanup-Mobile-Device-1205d2db


## Credits
Written by: Thomas Stensitzki / Sebastian Rubertus

## Social

* My Blog: http://justcantgetenough.granikos.eu
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:  http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de