# Manage-FolderPermissions GUI Tool

## Author:  
Agustin Gallegos

### Previous link:
<u>https://gallery.technet.microsoft.com/Manage-Folder-permissions-b5295673</u>  

## Examples  
### Example 1  
````powershell
PS C:\> .\Manage-FolderPermissionsGUI.ps1
````
In this exmaple the script will run, and will ask for a global admin credentials.  

## Version History:  
### 4.30 - 05/11/2020
 - Updated tool to connect to Exchange Online using new EXO v2 module.
### 4.20 - 02/19/2018
 - Changed On-premises connection logic, allowing the user to type the complete URL, and setting Authentication option.
### 4.00 - 06/04/2017
 - Optimized Import-PSSession command to import only necessary commands, and speed up session process.
 - Corrected drop-down menus to 2 separate variables.
 - Added colorized powershell lines, for better recognizing the operations performed.
 - changed hosts lines with time stamp and better explanation of the operation including user's DisplayName and folder.
 - Added transcript into the user's desktop to have detailed logging.
### 3.00 - 06/02/2017
 - Fixed lines between 164 and 172. In which output is only 1 liner, and needs to be added a single line to array.
 - Added an "APPLY TO ALL FOLDERS" combo entry. If you need to add or remove a user's permission to all folders, can apply this option.
### 2.00 - 04/19/2016
 - Initial Public Release.
### 1.00 - 03/31/2016
 - Project Start.