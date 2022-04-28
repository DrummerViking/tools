# Replace Rooms in Meetings  

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER RoomsCSVFilePath  
Sets the Rooms mapping file path. This file should have 2 columns named "PreviousRoom","newRoom".  

### PARAMETER MailboxesCSVFilePath  
Sets the users file path. This file should have 1 column named "PrimarySMTPAddress".  

### PARAMETER StartDate  
Sets the start date to look for meeting item in the user mailboxes. By default is the current date.  

### PARAMETER EndDate  
Sets the end date to look for meeting item in the user mailboxes. By default is 1 year after the current date.  

### PARAMETER ValidateUsersExistence  
If this Switch parameter is used, the script will not only connect using EWS, but will attempt to connect to EXO Powershell module and validate the users exists as valid mailboxes in Exchange Online.  

### PARAMETER ValidateRoomsExistence  
If this Switch parameter is used, the script will not only connect using EWS, but will attempt to connect to EXO Powershell module and validate the room mailboxes exists as valid recipients in Exchange Online.  

### PARAMETER EnableTranscript  
If this Switch parameter is used, all information displayed in the Powershell console, will be exported to the transcript file usually saved in "Documents" folder.  

## How to use the mapping file  
When composing the CSV for the rooms mailboxes mapping file, the script expects 2 columns.  
First column is the current Room mailbox being used in the meeting, and second column is the new room mailbox that will replace the one on the left.  
For example:  
| PreviousRoom | NewRoom |
|--------------|---------|
|RoomA@contoso.com | NewRoomReplacingA@contoso.com |
|RoomB@contoso.com | NewRoomReplacingB@contoso.com |
|RoomC@contoso.com | NewRoomReplacingC@contoso.com |

## Examples  
### Example 1  
```powershell
PS C:\> .\Replace-RoomsInMeetings.ps1 -ValidateRoomExistence
```
In this example the script will pop-up and prompt for the CSV with the mapping file for room accounts, and the CSV file for the users where to replace the rooms.  
Aside of connecting to EWS, the script will connect to EXO Powershell (it might ask for credentials again) and validate the rooms detailed in the mapping file exists as recipients in EXO.  
the script will look for meeting items since the current day and 1 year forward.  

### Example 2  
```powershell
PS C:\> .\Replace-RoomsInMeetings.ps1 -RoomsCSVFilePath C:\Temp\RoomsMappingFile.csv
```
In this example the script reads the Rooms mapping file from "C:\Temp\RoomsMappingFile.csv".  
Then will pop-up and prompt for the CSV file for the users where to replace the rooms.  
the script will look for meeting items since the current day and 1 year forward.  

### Example 3  
```powershell
PS C:\> .\Replace-RoomsInMeetings.ps1 -RoomsCSVFilePath C:\Temp\RoomsMappingFile.csv -MailboxesCSVFilePath C:\Temp\Users.Csv -EndDate 01/01/2025
```
In this example the script reads the Rooms mapping file from "C:\Temp\RoomsMappingFile.csv" and user's list from "C:\Temp\Users.Csv".  
the script will look for meeting items since the current day through January 1st 2025.  


## Version History:
### 1.04 - 04/28/2022
 - Fixed: Fixed recurrent meetings handling. The script now should be updating the whole series.  
### 1.03 - 04/18/2022
 - Fixed: Added logic to skip recurrent meetings. Still need to find some handling of recurrent meetings.
 - Added: Added Verbose output with more details.
### 1.01 - 04/04/2022
 - Fixed: Fixed dual connection to EWS and EXO if selecting Switch parameter "ValidateRoomExistence".
 - Added: Added functionally to validate user mailboxes as well with new Switch parameter "ValidateUsersExistence".
### 1.00 - 04/01/2022
 - First Release.
### 1.00 - 03/31/2022
 - Project start.