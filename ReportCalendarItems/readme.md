# Report Calendar items Tool

## Author:  
Agustin Gallegos  

## Examples  
### Example 1  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -EnableTranscript
```
In this exmaple the script will run, and will ask for a global admin credentials.  
It will pop out, asking for the CSV file containing the mailboxes to be read.  
The resultant file will be generated in the user's desktop.  
It will create a transcript file.  

### Example 2  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -EnableTranscript
```
In this example the script will run, and will ask for a global admin credentials.  
It will already use the CSV file "D:\Temp\rooms.csv" for the mailboxes to be read.  
The resultant file will be generated in the user's desktop.  

### Example 3  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -DestinationFolderPath "C:\Reports" -EnableTranscript
```
In this exmaple the script will run, and will ask for a global admin credentials.  
It will already use the CSV file "D:\Temp\rooms.csv" for the mailboxes to be read.  
The resultant file will be generated in the "C:\Reports" folder.  
It will create a transcript file.  

## Version History:
### 2.00 - 05/11/2020
 - Updated tool to connect to Exchange Online using oauth authentication. 
### 1.00 - 01/04/2019
 - First Release
### 1.00 - 01/04/2019
 - Project start