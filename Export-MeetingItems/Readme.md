# Export Meeting Items

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER ClientID
This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

### PARAMETER TenantID
This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

### PARAMETER ClientSecret
This is an optional parameter. String parameter with the Client Secret which is configured in the AzureAD App.  

### PARAMETER Mailboxes
This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.

### PARAMETER StartDate
This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider 1 year backwards from the current date.  

### PARAMETER EndDate
This is an optional parameter. The script will search for meeting items ending based on this EndDate backwards. If this parameter is ommitted, by default will consider 1 year forwards from the current date.

### PARAMETER ExportFolderPath
Insert target folder path named like "C:\Temp". By default this will be "$home\desktop"

### PARAMETER EnableTranscript
This is an optional parameter. Enable this parameter to write a powershell transcript in your 'Documents' folder.

## Examples  
### Example 1  
```powershell
PS C:\> .\Export-MeetingItems.ps1 -Mailboxes "user1@contoso.com" -EnableTranscript
```  
The script will ask for a user credential with impersonation permissions granted.  
will run against the "user1@contoso.com" mailbox and archive (if exists).  
Will Export meeting items to the default folder "$home\desktop" in a file named by 'user1-CalendaritemsReport.csv"  

### Example 2  
```powershell
# Following line requires to be connected to Exchange Online
PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
PS C:\> .\Export-MeetingItems.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -ExportFolderPath "C:\Reports" -EnableTranscript
```
The script will collect all user's primary SMTP addresses from mailboxes belonging to "Staff" department (this command line would need to be connected to EXO Powershell).  
Will run against each mailbox and archive (if exists).  
Will Export meeting items to the selected folder "C:\Reports" in a file named by '_alias_>_-CalendaritemsReport.csv" for each user account.  

### Example 3
```powershell
PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "HR"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
PS C:\> .\Export-MeetingItems.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -ExportFolderPath "C:\Reports" -ClientID "12345678" -TenantID "abcdefg" -ClientSecret "a1b2c3d4!#$"
```
The script will collect all user's primary SMTP addresses from mailboxes belonging to "HR" department (this command line would need to be connected to EXO Powershell).  
Will connect using App-Only permission (this requires an AzureAD app and does not require Exchange Impersonation permissions. More info [here](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth).  
Will run against each mailbox and archive (if exists).  
Will Export meeting items to the selected folder "C:\Reports" in a file named by '_alias_-CalendaritemsReport.csv" for each user account.  


## Version History:
### 1.00 - 06/16/2022
 - First Release.
### 1.00 - 06/14/2022
 - Project start.