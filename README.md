# Exchange Powershell tools
Exchange Powershell tools mostly use in Exchange Online (Office 365)


## Search-GUI tool (Exchange On-prem)
Allows admins to Search, export and Delete content from mailboxes.  
Also allows to Get RecoverableItems and Restore items.  
**Update: This tool will only work in Exchange On-premises. It relies on the command "Search-Mailbox" which has been deprecated from Exchange Online.**  
[File](/search-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/search-gui/search-GUI.ps1)  

----

## Get MRM Statistics Tool (Exchange On-prem and EXO)  
Allow admins to check current MRM Statistics and info for users.  
App brings Current Retention Policy and Tags.  
Can get current Managed Folder Assistant Cycle Stats for primary and Archive Mailbox.  
Recently added a new button, to get statistics on messages received daily in the last 7 days.  
Button available to issue a "Start-ManagedFolderAssistant" on the account.  
Button available to get ManagedFolderAssistants logs from mailbox.  
[File](/Get-MRMStatistics/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Get-MRMStatistics/Get-MRMStatistics.ps1)  

----

## Get MRM Romaing XML Stream from Mailbox (Exchange On-prem and EXO)  
This scripts allows to inspect the MRM configuration message in a user's mailbox.  
Allows to see the PR_ROMAING_XMLSTREAM data, and returned as a text.  
It also allows to delete this message.  
[File](/Get-MRMRoamingXMLData/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Get-MRMRoamingXMLData/Get-MRMRoamingXMLData.ps1)  

----

## Merge SoftDeleted Mailboxes using a GUI (EXO)  
Automate the process to create a New-MailboxRestoreRequest and verify the progress of it.
It will allow to export SourceAccount's ProxyAddresses in case needs to be imported in the target account.   
Allows to select and combine if we involve Archive Mailboxes.  
[File](/MergeMailboxes-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/MergeMailboxes-gui/MergeMailboxes-GUI.ps1)  

----

## Online Mailbox and Archive report using a GUI  (EXO)  
Get reports for Mailboxes and Archives hosted in Exchange Online.  
Report can be viewed live in powershell interface, or send as HTML report by email.  
[File](/OnlineArchiveReport-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/OnlineArchiveReport-gui/OnlineArchiveReport-GUI.ps1)  

----

## Delete Meetings using a GUI  (Exchange On-prem and EXO)  
Delete Meeting items from attendees when Organizers already left the company, in Exchange Online.  
You can pass a list of users/room mailboxes, and delete all meetings found from a specific Organizer.  

As this uses EWS, you will need a "master account" with Impersonation permissions. You can run:  
``` powershell
New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:<Account>
```  
This will give Impersonation permissions over all mailboxes in the tenant, so is critical that this account is not shared.  
We recommend that y ou use the tool initially for a single user/room mailbox, and check you have the expected experience.  
Is not recommended to run against an Organizer Mailbox. There is already a builtin command in Exchange Online service for this: `Remove-CalendarEvents`  
[File](/DeleteMeetings-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/DeleteMeetings-gui/DeleteMeetings-GUI.ps1)  

----  

## Replace Room locations in meetings for new ones (EXO)  
There are times that organizations needs to delete some Room Mailboxes, but if those mailboxes have some meeting items already scheduled, we actually need to replace for a new Room Mailbox.  
This is usually a tedious task that every meeting Organizer should do, by removing the previous Room mailbox, add the new one, and send the update to all attendees.  
We have made this script in order ease this task for admins.  

As this uses EWS, you will need a "master account" with Impersonation permissions on your mailboxes. You can run:  
``` powershell
New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:<Account>
```  
This will give Impersonation permissions over all mailboxes in the tenant, so is critical that this account is not shared.  

More info and details here:  
[File](/Replace-RoomsInMeetings/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Replace-RoomsInMeetings/Replace-RoomsInMeetings.ps1)  

----  

## Report Calendar Items  (EXO)  
Reports how many calendar items, per calendar year, some user/room mailboxes have in Exchange Online.  
Report can be exported to a DestinationFolderPath or by default to user's Desktop.  


As this uses EWS, you will need a "master account" with Impersonation permissions. You can run:  
``` powershell
New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:<Account>
```  
This will give Impersonation permissions over all mailboxes in the tenant, so is critical that this account is not shared.  

The report exports the following columns:  
> Mailbox, Subject, Organizer, Start Time, End Time, Received Time  

[File](/ReportCalendarItems/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/ReportCalendarItems/Report-CalendarItems.ps1)  

----

## Manage Mobile Devices using a GUI  (EXO)  
Allows admins to manage mobile devices in Exchange Online with a simplified GUI, and 'allow' or 'block' them in bulk.  
[File](/Manage-MobileDevices/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Manage-MobileDevices/Manage-Mobiles-GUI.ps1)  

----

## Manage Folder Permisions for Admins using a GUI  (Exchange On-prem and EXO)  
This file loads a GUI (Powershell Forms) to allow an admin to manage their user's mailbox folder permissions. It allows to add, remove and get permissions.  
It has a simple logic to try to connect to on-premises environments automatically.  
It has been tested in Exchange 2013 and Office 365.  
[File](/Manage-FolderPermissions-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Manage-FolderPermissions-gui/Manage-FolderPermissionsGUI.ps1)  

----

## Manage UserPhoto using a GUI  (Exchange On-prem and EXO)  
Allow admins to upload user Photos to Exchange Online using a GUI.  
We grant the option to create a RBAC Role Group, with the minimum permissions to list mailboxes and manage UserPhotos. This is intended for a help desk assignment.  
[File](/Manage-UserPhoto-gui/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Manage-UserPhoto-gui/Manage-UserPhotoGUI.ps1)  

----

## Parse Incident reports emails  (Exchange On-prem and EXO)  

Have you ever had a folder in your mailbox with a bunch of Incident report emails? When the time comes to look for all the reports matching a rule, or a sender, you can't just look at your e-mails one by one.  
With this script, you will read all the e-mail reports in a folder, and extract that content to a CSV, so you can easily open with a spreadsheet editor and perform easier queries against it.  
The resultant file in the user's Desktop, will have the following columns:  
> Received Time, Report Id, Message Id, Sender, Subject, To, Rule Hit  

[File](/Parse-IncidentReports/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Parse-IncidentReports/Parse-IncidentReports.ps1)  

----

## Get Quarantine Report based on group members (EXO)  

Generate HTML report listing quarantine messages for a Security Group and its members.  
This report can be sent by e-mail to a list of recipients and/or tenant Admins and/or to the group member themselves.  
The report will be saved to the user's desktop by default or can be modified in the ReportFilePath parameter.  

[File](Get-QuarantineReport/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Get-QuarantineReport/Get-QuarantineReport.ps1)  

----

## Test Autodiscover V2  
Tests AutodiscoverV2 against Office365 or Exchange On-premises with different protocols available.  
[File](/Test-Autodiscover/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Test-Autodiscover/Test-Autodiscover.ps1)  

----

## Collect FreeBusy information  (Exchange On-prem and EXO)  
Collects info usually requested by Microsoft support to troubleshoot FreeBusy issues.  
[File](/CollectFBLogs/)

----

##  Collect SMTP Auth logs  (under construction)  
Collects info usually requested by Microsoft support to troubleshoot SMTP client submission issues.  
[File](/CollectSMTPLogs/)  

----  

## Get Exchange Server Info Tool  (Exchange On-prem)  

- This script checks the On-Premises Exchange servers for general information by getting:  
> Server's name, build number (and RU/CU name), AD Site and roles, .NET version, V++ versions.

- If you have a mixed environment running multiple versions of Exchange, it is recommended to run the script from your newest version available.  
- Includes a "CASLoad" switch, to collect some Protocol load counters, Netlogon's Semaphores, and check ASA credentials if it is running in a local CAS Server.  
- Includes an "Autodiscover" switch, to collect SCPs.  
[File](/Get-ExchangeServerInfo/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/tools/master/Get-ExchangeServerInfo/Get-ExchangeServerInfo.ps1)