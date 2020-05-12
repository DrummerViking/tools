# Get-MRMStatistics Tool

## Authors:  
Agustin Gallegos  
Nelson Riera  

## Version History:  
### 1.90 - 05/11/2020
 - Updated tool to connect to Exchange Online using new EXO v2 module.
### 1.80 - 10/08/2017
 - Fixed process 2, to get RetentionPolicyTags and details as it was not pulling SystemTags.
### 1.70 - 07/17/2017
 - Added "Get MFA logs" button to retrieve latest MFA logs for a mailbox
### 1.50 - 07/04/2017
 - Fixed process 2, to get RetentionPolicyTags and details.
 - Added Status Bar to Main Window
### 1.32 - 12/23/2016
 - Optimizing Import-PSSession command, to just import necessary cmdlets, and speed up session process
### 1.31 - 10/25/2016
 - Changed first popup message asking for endpoint, that only Global Admins can run this tool. User's credentials will not be able to run the commands in the background.
### 1.3  - 10/24/2016
 - Added Information message on Open. Added check for Mailbox existence, and validate Policy assigned.    
### 1.1  - 10/21/2016
 - First Release
###	1.00 - 10/16/2016
 - Project start