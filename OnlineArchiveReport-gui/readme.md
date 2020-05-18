# OnlineArchiveReport-GUI Tool  

## Authors:  
Agustin Gallegos  
Nelson Riera  

## Examples  
### Example 1  
```powershell
PS C:\> .\OnlineArchiveReport-GUI.ps1 -EnableTranscript
```
In this exmaple the script will run, and will ask for a global admin credentials.  
It will generate a transcript file.  

## Version History:
### 1.23 - 05/08/2020
 - Updated tool to connect using new EXO Powershell v2 module.
### 1.22 - 11/07/2018
 - Added a sleep between each mailbox loop, in order to prevent Powershell micro delays.
 - Added a switch paratemer to enable Powershell Transcript if desired.
### 1.20 - 10/15/2018
 - Added a progress status on each mailbox loop.
 - Added additional column calculating current percentage used of RecoverableItems quota.
### 1.00 - 10/09/2018
 - First Release.
### 1.00 - 10/09/2018
 - Project start.