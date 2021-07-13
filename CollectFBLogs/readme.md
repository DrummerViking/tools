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
  
  
## What does these scripts collects?  
  
### Collect-DAUTHTroubleshootinglogs.ps1 collects:  
#### On-premises data:  
- Federation Trust  
- Federated Organization Identifier  
- Organization Relationships  
- EWS Virtual Directories  
- Autodiscover Virtual Directories  
- Remote Mailbox info  
- On-premises Mailbox info  
- Tests Federation Trust  
- Tests Federation Trust Certificate  
- Availability Address Spaces  
- Sharing Policies  
- Receive Connectors  
- Send Connectors  

#### Exchange Online data:  
- Federation Trust  
- Federated Organization Identifier  
- Organization Relationships  
- Mail User info  
- Cloud's Mailbox info  
- Sharing Policies  
- Inbound Connectors  
- Outbound Connectors  
- Domain's Federation Information  

----

### CollectOAUTHTroubleshootinglogs.ps1 collects:  
#### On-premises data:  
- AuthServer info  
- PartnerApplication
- EWS Virtual Directories  
- Autodiscover Virtual Directories  
- IntraOrganizationConnectors
- Availability Address Spaces  
- Remote Mailbox info  
- On-premises Mailbox info  
- Tests OAUTH Connectivity to EWS service  
- Tests OAUTH Connectivity to AutoD service  
- Receive Connectors  
- Send Connectors  

#### Exchange Online data:  
- MSOL Service Principal Credentials  
- MSOL Service Principals  
- IntraOrganizationConnector  
- Mail User info  
- Cloud's Mailbox info  
- Inbound Connectors  
- Outbound Connectors  
- Tests OAUTH Connectivity to on-premises EWS service  
- Tests OAUTH Connectivity to on-premises AutoD service