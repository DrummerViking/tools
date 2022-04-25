# Inject mail items into the Archive's Inbox

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER TargetSMTPAddress  
Use this optional parameter to set an impersonation SMTP address.  
    
### PARAMETER SampleFileName  
File path to a sample file to be attach. If this parameter is ommitted, a test file of 34MB will be created.  

### PARAMETER NumberOfMessages  
This is the amount of messages to be created in the Public Folder. By default will attempt 100.
    
### PARAMETER EnableTranscript  
Use this Switch parameter to enable Powershell Transcript.  

### PARAMETER UseBasicAuth  
Use this Switch parameter to connect to EWS using Basic Auth. By default the script will attempt to connect using Modern Auth.  


## Examples  
### Example 1  
```powershell
PS C:\> Inject-MailItemsIntoArchive.ps1 -NumberOfMessages 10
```
The script will request the user credentials.  
Will attempt to inject 10 messages into the user Archive's Inbox.  

### Example 2  
```powershell
PS C:\> Inject-MailItemsIntoPF.ps1 -TargetSMTPAddress "impersonated@contoso.com" -EnableTranscript -UseBasicAuth
```
The script will request the user credentials which have impersonation permissions to open mailbox "impersonated@contoso.com".  
Will attempt to inject 100 messages (default value) into the user Archive's Inbox.  
Will save all powershell output to the Transcript file.  
Will connect to EWS using Basic Auth instead of Modern Auth.  

## Version History:  
### 1.00 - 04/20/2022
 - First Release.
### 1.00 - 04/20/2022
 - Project start.