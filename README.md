# Exchange Powershell tools
Exchange Powershell tools mostly use in Exchange Online (Office 365)

## Search-GUI tool
Allows admins to Search, Delete content from mailboxes.  
Also allows to Get RecoverableItems and Restore items.  
[File](/search-gui/)

----

## Merge SoftDeleted Mailboxes using a GUI
Automate the process to create a New-MailboxRestoreRequest and verify the progress of it.
It will allow to export SourceAccount's ProxyAddresses in case needs to be imported in the target account.   
Allows to select and combine if we involve Archive Mailboxes.  
[File](/MergeMailboxes-gui/)

----

## Online Mailbox and Archive report using a GUI  
Get reports for Mailboxes and Archives hosted in Exchange Online.  
Report can be viewed live in powershell interface, or send as HTML report by email.  
[File](/OnlineArchiveReport-gui/)

----

## Delete Meetings using a GUI  
Delete Meeting items from attendees when Organizers already left the company, in Exchange Online.  
[File](/DeleteMeetings-gui/)

----

## Report Calendar Items  
Reports how many calendar items, per calendar year, some user/room mailboxes have in Exchange Online.  
Report can be exported to a DestinationFolderPath or by default to user's Desktop.  
[File](/ReportCalendarItems/)

----

## Manage Mobile Devices using a GUI  
Allows admins to manage mobile devices in Exchange Online with a simplified GUI, and 'allow' or 'block' them in bulk.  
[File](/Manage-MobileDevices/)

----

## Collect FreeBusy information
Collects info usually requested by Microsoft support to troubleshoot FreeBusy issues.  
[File](/CollectFBLogs/)