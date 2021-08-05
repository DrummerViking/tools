# Get-MRMRoamingXMLData script  

## Authors:  
Agustin Gallegos  

## Info:  
Copy both .PS1 and .DLL files to the same folder in order to run the script.  

## Examples  
### Example 1  
```powershell
PS C:\> .\Get-MRMRoamingXMLData.ps1
```
In this example the script will ask for the user's credentials to be checked and get the MRM Roaming XML Data.  

### Example 2  
```powershell
PS C:\> .\Get-MRMRoamingXMLData.ps1 -DeleteConfigurationMessage
```
In this example the script will delete the 'IPM.Configuration.MRM' message from the user's mailbox.  
An Administrator should run `Start-ManagedFolderAssistant` to issue MRM service and recreate the message. 

### Example 3  
```powershell
PS C:\> .\Get-MRMRoamingXMLData.ps1 -TargetSMTPAddress 'anotherUser@domain.com'
```
In this example the script will ask for the Admin's credentials to authenticate. And will actually open 'anotherUser@domain.com' mailbox to check and get the MRM Roaming XML Data.  
 

## Version History:  
### 2.30 - 11/10/2020
 - Added 'TargetSMTPAddress' in order to open another user's mailbox with Impersonation permissions.
### 2.20 - 05/25/2020
 - Added 'DeleteConfigurationMessage' Switch parameter, to delete the IPM.Configuration.MRM message.
 - Added some try/catch blocks to catch error messages and properly show to the user.
### 2.00 - 05/14/2020
 - Updated tool to connect to Exchange Online using oauth authentication.
### 1.00 - 08/22/2018
 - Project Start
### 1.00 - 08/22/2018
 - First Release
