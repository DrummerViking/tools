# Test Autodiscover

## Author:  
Agustin Gallegos  

## Summary  
Test AutodiscoverV2 either against Office365 ( default server name ) or Exchange On-premises.  
You can select one of the available protocols:  
> "AutodiscoverV2","ActiveSync","Ews","Rest","Substrate","SubstrateNotificationService","SubstrateSearchService","OutlookMeetingScheduler".  

## Examples  
### Example 1  
```powershell
PS C:\> Test-Autodiscover -EmailAddress onpremUser@contoso.com -Protocol AutodiscoverV2
```
In this example it will show the autodiscover URL for the onpremises user, queried against outlook.office365.com  

### Example 2  
```powershell
PS C:\> Test-Autodiscover -EmailAddress cloudUser@contoso.com -Protocol EWS -Server mail.contoso.com
```
In this example it will show the EWS URL for the cloud user, queried against an on-premises endpoint 'mail.contoso.com'.  

## Version History: 
### 1.00 - 10/11/2020
 - First Release
### 1.00 - 10/11/2020
 - Project start