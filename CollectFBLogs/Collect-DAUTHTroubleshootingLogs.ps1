#Requires -RunAsAdministrator
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
[CmdletBinding()]
Param(
    [Parameter(Position=1,Mandatory = $False,HelpMessage = 'Primary Email Address of an on-premises mailbox you want to check...')]
    [string]$OnpremisesUser = '',

    [Parameter(Position=2,Mandatory = $False,HelpMessage = 'Primary Email Address of a cloud mailbox you want to check...')]
    [string]$CloudUser = ''
)

# Disclaimer
$disclaimer = @"
#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
#  
#################################################################################
"@
Write-Host $disclaimer -foregroundColor Yellow
Write-Host " " 
if ( !(Get-Module PSFramework) -and !(Get-Module PSFramework -ListAvailable) )
{
    Install-Module PSFramework -Force
}

#using C:\TEMP\MSlogs folder
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Checking C:\TEMP\MSlogs Folder"
$folder = "C:\TEMP\MSlogs" 
if (-not (Test-path $folder) )
{ 
    #Create the directory 
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Creating Directory $folder" -DefaultColor Green 
    $null = [System.IO.Directory]::CreateDirectory($folder) 
} 
else
{ 
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] $folder is already created!" -DefaultColor Yellow 
}
Set-Location C:\temp\MSlogs

#setting variables
if($OnpremisesUser -eq ''){ $OnpremisesUser = Read-Host -Prompt "please enter the Primary Email Address of an on-premises mailbox" }
if($CloudUser -eq '')     { $CloudUser = Read-Host -Prompt "please enter the Primary Email Address of a cloud mailbox" }

#---------------------------
# On-premises side
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Using Exchange On-premises Powershell"

$ts = Get-Date -Format "yyyy-MM-dd hh_mm_ss" 
$FormatEnumerationLimit = -1
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Federation Trust"
Get-FederationTrust | export-clixml -path "$ts.OnPrem_FedTrust.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Federated Organization Identifier"
Get-FederatedOrganizationIdentifier | export-clixml -path "$ts.OnPrem_FedOrgID.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Organization Relationships"
Get-OrganizationRelationship | export-clixml -path "$ts.OnPrem_OrgRel.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting EWS Virtual Directories"
Get-WebServicesVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_EWSVdir.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Autodiscover Virtual Directories"
Get-AutoDiscoverVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_AutodVdir.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Remote Mailbox info"
Get-RemoteMailbox $CloudUser | export-clixml -path "$ts.OnPrem_RemoteMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting On-premises Mailbox info"
Get-Mailbox $OnpremisesUser | export-clixml -path "$ts.OnPrem_OnPremisesMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] testing Federation Trust"
Test-FederationTrust -UserIdentity $OnpremisesUser -Verbose | export-clixml -path "$ts.OnPrem_TestFedTrust.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] testing Federation Trust Certificate"
Test-FederationTrustCertificate | export-clixml -path "$ts.OnPrem_TestFedTrustCert.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Availability Address Spaces"
Get-AvailabilityAddressSpace | export-clixml -path "$ts.OnPrem_AvailAddrSpc.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Sharing Policies"
Get-SharingPolicy | export-clixml -path "$ts.OnPrem_SharingPolicy.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Receive Connectors"
Get-ReceiveConnector | export-clixml -path "$ts.OnPrem_ReceiveConnectors.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Send Connectors"
Get-SendConnector | export-clixml -path "$ts.OnPrem_SendConnectors.xml"

##---------------------------
#connecting to Cloud side
Write-Host ""
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online Powershell"

$LiveCred = Get-Credential -Message "Please enter your Global Admin Credentials"
if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) )
{
    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $LiveCred -Prefix EO

$FormatEnumerationLimit = -1
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Federation Trust"
Get-EOFederationTrust | export-clixml -path "$ts.Cloud_FedTrust.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Federated organization Identifier"
Get-EOFederatedOrganizationIdentifier | export-clixml -path "$ts.Cloud_FedOrgID.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Organization Relationships"
Get-EOOrganizationRelationship | export-clixml -path "$ts.Cloud_OrgRel.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Mail User info"
Get-EOMailUser $OnpremisesUser | export-clixml -path "$ts.Cloud_OnPremisesMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Mailbox info"
Get-EOMailbox $CloudUser | export-clixml -path "$ts.Cloud_MBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Sharing policies"
Get-EOSharingPolicy | export-clixml -path "$ts.Cloud_SharingPolicy.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Inbound Connectors"
Get-EOInboundConnector | export-clixml -path "$ts.Cloud_InboundConnectors.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Outbound Connectors"
Get-EOOutboundConnector | export-clixml -path "$ts.Cloud_OutboundConnectors.xml"

#getting Federationinformation from on-premises user's Primary Email Address Domain
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting domain's Federation Information"
$domain = $OnpremisesUser.Substring($OnpremisesUser.IndexOf('@')+1)
Get-EOFederationInformation -Domainname $domain | export-clixml -path "$ts.Cloud_FedInfo.xml"

#---------------------------
# Disconnecting from Cloud side

Write-Host ""
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Disconnecting from Exchange Online Powershell"
Disconnect-ExchangeOnline -Confirm:$False

# compressing log files
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting files"
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Using 'C:\Temp\MSLogs\loggingFiles.zip'"
$logzipfile = 'C:\Temp\MSLogs\loggingFiles.zip'
if ( Test-Path $logzipfile )
{
    remove-item $logzipfile
}
if ( $PSVersionTable.PSVersion.Major -lt 5 )
{
    try
    {
        Add-Type -assembly "system.io.compression.filesystem"
        [io.compression.zipfile]::CreateFromDirectory($folder, $logzipfile) 

        #deleting XML files
        Get-ChildItem $folder\*.xml | remove-item -Force
        Get-ChildItem $folder\*.cer | remove-item -Force
    }
    catch {    }

} 
else
{
    Compress-Archive -Path $folder\* -CompressionLevel Fastest -DestinationPath $logzipfile
    
    #deleting XML files
    Get-ChildItem $folder\*.xml | remove-item -Force
    Get-ChildItem $folder\*.cer | remove-item -Force
}
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Please collect 'C:\Temp\MSLogs\loggingFiles.zip' and send it to your support engineer."