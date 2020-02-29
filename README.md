# Remove-MobileDevicePartnership

Remove Exchange Server 2013+ Mobile Device Partnerships 

## Description

This script removes mobile device association from user mailboxes that have been inactive for more than X days.

Use the settings.xml to configure your email server settings and the min number of days for old mobile devices.

## Requirements

- Exchange Server 2013, 2016, 2019
- Windows Server 2012 R2+
- Exchange Management Shell

## Parameters

### MailboxFilter

Check only a give mailbox or some mailboxes for aged mobile devices. Preferably, you use the Alias. This parameter works with wildcards, e.g., USER*

### SendMail

Send the list of found mobile devices by email. Email settings are controlled by a dedicated settings.xml file. See script link for more details.

### ReportOnly

Just create a report for all found mobile devices, but DO NOT DELETE the mobile device partnerships.

## Examples

``` PowerShell
.\Remove-MobileDevicePartnership.ps1
```

Remove old mobile device partnerships without sending a report email

``` PowerShell
.\Remove-MobileDevicePartnership.ps1 -SendMail
```

Remove old mobile device partnerships and send a report email

``` PowerShell
.\Remove-MobileDevicePartnership.ps1 -ReportOnly
```

Search for old mobile device partnerships and write results as CSV to disk

``` PowerShell
.\Remove-MobileDevicePartnership.ps1 -MailboxFilter USERALIAS -SendMail
```

Remove old mobile device partnerships for a single mailbox and send a report email

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Download and vote at TechNet Gallery

- [https://gallery.technet.microsoft.com/Cleanup-Mobile-Device-1205d2db](https://gallery.technet.microsoft.com/Cleanup-Mobile-Device-1205d2db)

## Credits

Written by: Thomas Stensitzki

Stay connected:

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn:	[http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

Additonal Credits

- Sebastian Rubertus developed the original version for Exchange Server 2010.