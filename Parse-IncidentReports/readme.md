# Parse incident Report emails to CSV  

## Author:  
Agustin Gallegos  

## Examples  
### Example 1  
```powershell
PS C:\> .\Parse-IncidentReports.ps1 -OrgAdmins
```
In this exmaple the script will run, and will ask for a global admin credentials.  
Once the report is ready, it will send it to all Global Admins found in the tenant.  

### Example 2  
````powershell
PS C:\> .\Parse-IncidentReports.ps1 -OrgAdmins -Recipients "ExternalAuditing@Audits.com"
````
In this exmaple the script will run, and will ask for a global admin credentials.  
Once the report is ready, it will send it to all Global Admins found in the tenant as well as to the external account "ExternalAuditing@Audits.com".  

## Parameters list  

### PARAMETER Recipients  
The email address you want the report to be sent to.  

### PARAMETER OrgAdmins  
Send report to Organization Admins detected.  

## Version History:  
### 1.80 - 11/09/2022  
 - Updated logic to list and select the folder where incident reports are located.
 - Updated command to connect to EXO if we need to fetch OrgAdmins
### 1.60 - 05/11/2020
 - Updated tool to connect to Exchange Online using oauth authentication.
### 1.50 - 06/19/2017
 - Added Email report to "Recipients" parameter
 - Added "OrgAdmins" switch, that can be combined to send the emails to specific recipients and/or Orgadmins
### 1.10 - 06/08/2017
 - Added "Received Time" column to resultant CSV file.
### 1.00 - 03/21/2017
 - Initial Public Release.
 