# DeleteMeetings-GUI Tool  

## Author:  
Agustin Gallegos  

## Info:  
Copy both .PS1 and .DLL files to the same folder in order to run the script.  

## Examples:  
### Example 1  
```powershell
PS C:\> .\DeleteMeetings-GUI.ps1 -EnableTranscript
```
In this example the script will run and create a transcript file. It will log the exported list of items to the user's desktop. 

### Example 2  
```powershell
PS C:\> .\DeleteMeetings-GUI.ps1 -LogFolder "C:\Temp"
```
In this example the script will run and it will log the exported list of items to the 'C:\Temp' folder.  

### PARAMETER EnableTranscript  
Enable this parameter to write a powershell transcript in your 'Documents' folder.  

### PARAMETER LogFolder  
Sets the folder to export the logs generated. If this parameter is omitted, logs will be generated in the user's Desktop.  

## Version History:  
### 2.01 - 09/03/2021  
 - Updated tool and add an optional parameter to export log files to a custom folder.  
### 2.00 - 05/11/2020
 - Updated tool to connect to Exchange Online using oauth authentication.
### 1.30 - 01/03/2019
 - Remove hardcoded timeframe of 180 days. Now user can select desired time frame, including past items
 - Added 'Subject' column to results. **Take into account, EXO's default Calendar Processing is to Delete the Subject for Room Mailboxes**
### 1.00 - 12/27/2018
 - First Release
### 1.00 - 12/27/2018
 - Project start
