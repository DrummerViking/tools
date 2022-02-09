# Get QuarantineReport Tool  

## Authors:  
Agustin Gallegos  

## Examples  
### Example 1  
```powershell
PS C:\> .\Get-QuarantineReport.ps1 -GroupAddress InfoSecurity -Recipients "externalUser@Outlook.com"
```
In this example the script will get all members from group "InfoSecurity".  
Will get all quarantine messages related to these recipients.  
It will ask for a global admin credentials.  
It will generate the report file and save it to the user's Desktop.  
Later will send the report by email to the recipient "externalUser@Outlook.com".  

### Example 2  
```powershell
PS C:\> .\Get-QuarantineReport.ps1 -GroupAddress VIPUsers -OrgAdmins -ReportFilePath "C:\Temp\Quarantine Report.html"
```
In this example the script will get all members from group "VIPUsers".  
Will get all quarantine messages related to these recipients.  
It will ask for a global admin credentials.  
It will generate the report file and save it to "C:\Temp\Quarantine Report.html".  
Later will send the report by email to the tenant's Global Admin group.  

### PARAMETER GroupAddress  
Group Alias you want to get the list of members of

### PARAMETER Recipients
comma separated list of recipients to which the report should be sent to.

### PARAMETER OrgAdmins
This is a switch Parameter. Using it, will send the report to every Global Admin in the tenant. Can be combined together with "recipients" parameter.

### PARAMETER EmailtoGroupMembers
This is a switch Parameter. Using it, will send the report to the group you are collecting the report for.

### PARAMETER ReportFilePath
Path where the HTML report will be saved. By default will be in the user's desktop named "Quarantine report.html".  

## Version History:
### 2.00 - 02/09/2022
 - Updated script to Github repository.
 - Updated script to use new EXO PS module.
 - Added additional modules dependency.
### 1.00 - 03/02/2018
 - First Release.
### 1.00 - 03/02/2018
 - Project start.