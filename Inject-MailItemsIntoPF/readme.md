# Inject mail items into a Public Folder  

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER TargetPublicFolder  
This is the path to the public folder. It must start with a backslash. For example "\Company Root\folder1\SubFolder2". It should not end with a backslash neither. This is a mandatory parameter.  
    
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
PS C:\> Inject-MailItemsIntoPF.ps1 -TargetPublicFolder "\My Company Root\folder1" -NumberOfMessages 10
```
The script will request the user credentials.  
Will validate the folder path exists.  
Will attempt to inject 10 messages into the target folder "\My Company Root\folder1".  

### Example 2  
```powershell
PS C:\> Inject-MailItemsIntoPF.ps1 -TargetPublicFolder "\Corp\subfolder2" -EnableTranscript -UseBasicAuth
```
The script will request the user credentials.  
Will validate the folder path exists.  
Will attempt to inject 100 messages (default value) into the target folder.  
Will save all powershell output to Transcript file.  


## Version History:  
### 1.00 - 04/20/2022
 - First Release.
### 1.00 - 04/20/2022
 - Project start.