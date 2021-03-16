# Get-ExchangeServerInfo Tool  

## Author:  
Agustin Gallegos  

## Examples  
### Example 1  
```powershell
PS C:\> .\Get-ExchangeServerInfo.ps1 -Server "EXCHSERVER1"
```
Runs against a single Exchange server.  

### Example 2  
```powershell
PS C:\> .\Get-ExchangeServerInfo.ps1 -Server "EXCHSERVER1" -CASLoad
```
Runs against a single Exchange server including CAS related information.  

### Example 3  
```powershell
PS C:\> .\Get-ExchangeServerInfo.ps1 -Site "Site1"
```
Runs against all servers located in the specified Site.  

## Version History:  
### 1.105 - 03/16/2020
 - Updated build versions for March 2021 releases
### 1.105 - 12/17/2020
 - Updated build versions for December 2020 releases
### 1.102 - 09/29/2020
 - Updated build versions for September 2020 releases
### 1.100 - 06/17/2020
 - Updated build versions for June 2020 releases
### 1.97 - 03/30/2020
 - Updated build versions for March 2020 releases
### 1.95 - 12/17/2019
 - Updated build versions for December 2019 releases
### 1.93 - 10/01/2019
 - Updated build versions for September 2019 releases
### 1.91 - 09/16/2019
 - Updated build versions for June 2019 releases
### 1.87 - 02/12/2019
 - Updated build versions for February 2019 releases
### 1.83 - 10/16/2018
 - Updated build version for October 2018 releases.
### 1.82 - 10/09/2018
 - Added E2010 SP3 RU23 and RU24 build numbers.
### 1.80 - 06/19/2018
 - updated build versions for June 2018 releases.
 - Added Column to check if Visual C++ 2013 Redistributable is installed or not
### 1.72 - 03/22/2018
 - updated build versions for March 2018 releases.
### 1.70 - 02/20/2018
 - Added .NET Framework version column
### 1.62 - 02/14/2018
 - updated build versions for December 2017 releases.
### 1.59 - 09/20/2017
 - updated build versions for September 2017 releases.
### 1.57 - 06/29/2017
 - updated build versions for June 2017 releases.
### 1.55 - 03/23/2017
 - updated build versions for March 2017 releases.
### 1.52 - 12/14/2016
 - updated build versions for December 2016 releases.
### 1.48 - 10/13/2016
 - updated build versions for September 2016 releases.
### 1.47 - 07/01/2016
 - updated build versions for June 2016 releases. Also added an "Autodiscover" switch to bring Autodiscover related info
### 1.43 - 03/17/2016
 - Corrected "CheckASACredentials" switch to skip if running in Exchange 2007 server.
### 1.40 - 03/15/2016
 - updated build versions for March 2016 releases. Also moved "CheckASACredentials" to a separate Switch
### 1.20 - 02/16/2016
 - Initial Public Release.
### 1.00 - 01/28/2016
 - Project Start.