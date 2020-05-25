# Collect FreeBusy information
Collects info usually requested by Microsoft support to troubleshoot FreeBusy issues.  
Please run these scripts from an Exchange On-premises Powershell Shell. The commands will use the existing session to get on-premises data, and will connect to EXO to get cloud's data.  

## Authors:  
Agustin Gallegos

## Examples  
### Example 1  
```powershell
PS C:\> .\Collect-DAUTHTroubleshootinglogs.ps1 -OnpremisesUser "onpremuser@contoso.com" -CloudUser "clouduser@contoso.com"
```
In this example will collect DAUTH relevant logs with two sample users data.  

### Example 2  
```powershell
PS C:\> .\Collect-OAUTHTroubleshootinglogs.ps1 -OnpremisesUser "onpremuser@contoso.com" -CloudUser "clouduser@contoso.com"
```
In this example will collect OAUTH relevant logs with two sample users data.
