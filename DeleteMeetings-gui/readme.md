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
In this example the script will run and create a transcript file.  

## Version History:
### 2.00 - 05/11/2020
 - Updated tool to connect to Exchange Online using oauth authentication.
### 1.30 - 01/03/2019
 - Remove hardcoded timeframe of 180 days. Now user can select desired time frame, including past items
 - Added 'Subject' column to results. **Take into account, EXO's default Calendar Processing is to Delete the Subject for Room Mailboxes**
### 1.00 - 12/27/2018
 - First Release
### 1.00 - 12/27/2018
 - Project start
