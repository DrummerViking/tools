# Report Calendar items Tool

## Author:  
Agustin Gallegos  

## Info:  
Copy both .PS1 and .DLL files to the same folder in order to run the script.  

## Examples  
### Example 1  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -EnableTranscript
```
In this example the script will run, and will ask for a global admin credentials.  
It will pop out, asking for the CSV file containing the mailboxes to be read.  
The resultant file will be generated in the user's desktop.  
It will create a transcript file.  

### Example 2  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -EnableTranscript
```
In this example the script will run, and will ask for a global admin credentials.  
It will use the CSV file "D:\Temp\rooms.csv" for the mailboxes to be read.  
The resultant file will be generated in the user's desktop.  

### Example 3  
```powershell
PS C:\> .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -DestinationFolderPath "C:\Reports" -EnableTranscript
```
In this example the script will run, and will ask for a global admin credentials.  
It will use the CSV file "D:\Temp\rooms.csv" for the mailboxes to be read.  
The resultant file will be generated in the "C:\Reports" folder.  
It will create a transcript file.  

## Parameters list  

### PARAMETER EnableTranscript  
Enable this parameter to write a powershell transcript in your 'Documents' folder.  

### PARAMETER CSVFile  
CSV file must contain a unique header named "PrimarySMTPAddress".  

### PARAMETER DestinationFolderPath  
Insert target folder path name like "C:\Temp".  

## Version History:
### 2.00 - 05/11/2020
 - Updated tool to connect to Exchange Online using oauth authentication. 
### 1.00 - 01/04/2019
 - First Release
### 1.00 - 01/04/2019
 - Project start