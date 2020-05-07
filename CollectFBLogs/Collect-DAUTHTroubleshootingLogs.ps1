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

#using C:\TEMP\MSlogs folder
    Write-Host "Checking C:\TEMP\MSlogs Folder" -ForegroundColor Yellow
    $folder = "C:\TEMP\MSlogs" 
    if (-not (Test-path $folder) ) 
    { 
        #Create the directory 
        Write-Host "Creating Directory $folder" -ForegroundColor Green 
        $null = [System.IO.Directory]::CreateDirectory($folder) 
    }
    else
    { 
        Write-Host "$folder is already created!" -ForegroundColor Yellow 
    } 
    Set-Location C:\temp\MSlogs

#setting variables
if($OnpremisesUser -eq ''){ $OnpremisesUser = Read-Host -Prompt "please enter the Primary Email Address of an on-premises mailbox" }
if($CloudUser -eq '')     { $CloudUser = Read-Host -Prompt "please enter the Primary Email Address of a cloud mailbox" }

#---------------------------
# On-premises side
Write-Host "Using Exchange On-premises Powershell" -ForegroundColor Cyan 

$ts = Get-Date -Format "yyyy-MM-dd hh_mm_ss" 
$FormatEnumerationLimit = -1
Get-FederationTrust | export-clixml -path "$ts.OnPrem_FedTrust.xml"
Get-FederatedOrganizationIdentifier | export-clixml -path "$ts.OnPrem_FedOrgID.xml"
Get-OrganizationRelationship | export-clixml -path "$ts.OnPrem_OrgRel.xml"
Get-WebServicesVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_EWSVdir.xml"
Get-AutoDiscoverVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_AutodVdir.xml"
Get-RemoteMailbox $CloudUser | export-clixml -path "$ts.OnPrem_RemoteMBX.xml"
Get-Mailbox $OnpremisesUser | export-clixml -path "$ts.OnPrem_OnPremisesMBX.xml"
Test-FederationTrust -UserIdentity $OnpremisesUser -Verbose | export-clixml -path "$ts.OnPrem_TestFedTrust.xml"
Test-FederationTrustCertificate | export-clixml -path "$ts.OnPrem_TestFedTrustCert.xml"
Get-AvailabilityAddressSpace | export-clixml -path "$ts.OnPrem_AvailAddrSpc.xml"
Get-SharingPolicy | export-clixml -path "$ts.OnPrem_SharingPolicy.xml"
Get-ReceiveConnector | export-clixml -path "$ts.OnPrem_ReceiveConnectors.xml"
Get-SendConnector | export-clixml -path "$ts.OnPrem_SendConnectors.xml"

#---------------------------
#connecting to Cloud side
Write-Host ""
Write-Host "Connecting to Exchange Online Powershell" -ForegroundColor Cyan

$LiveCred = Get-Credential -Message "Please enter your Global Admin Credentials"
            
if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) )
{
    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $LiveCred -Prefix EO

$FormatEnumerationLimit = -1

Get-EOFederationTrust | export-clixml -path "$ts.Cloud_FedTrust.xml"
Get-EOFederatedOrganizationIdentifier | export-clixml -path "$ts.Cloud_FedOrgID.xml"
Get-EOOrganizationRelationship | export-clixml -path "$ts.Cloud_OrgRel.xml"
Get-EOMailUser $OnpremisesUser | export-clixml -path "$ts.Cloud_OnPremisesMBX.xml"
Get-EOMailbox $CloudUser | export-clixml -path "$ts.Cloud_MBX.xml"
Get-EOAvailabilityAddressSpace | export-clixml -path "$ts.Cloud_AvailAddrSpc.xml"
Get-EOSharingPolicy | export-clixml -path "$ts.Cloud_SharingPolicy.xml"
Get-EOInboundConnector | export-clixml -path "$ts.Cloud_InboundConnectors.xml"
Get-EOOutboundConnector | export-clixml -path "$ts.Cloud_OutboundConnectors.xml"

#getting Federationinformation from on-premises user's Primary Email Address Domain
$domain = $OnpremisesUser.Substring($OnpremisesUser.IndexOf('@')+1)
Get-EOFederationInformation -Domainname $domain | export-clixml -path "$ts.Cloud_FedInfo.xml"

#---------------------------
#Disconnecting from Cloud side
Write-Host ""
Write-Host "Disconnecting from Exchange Online Powershell" -ForegroundColor Cyan
Disconnect-ExchangeOnline

# compressing log files
$logzipfile = 'C:\Temp\MSLogs\loggingFiles.zip'
if( Test-Path $logzipfile )
{
    remove-item $logzipfile -Force
}
if( $PSVersionTable.PSVersion.Major -lt 5 )
{
    try
    {
        Add-Type -assembly "system.io.compression.filesystem"
        [io.compression.zipfile]::CreateFromDirectory($folder, $logzipfile) 

        #deleting XML files
        Get-ChildItem $folder\*.xml | remove-item -Force
    }
    catch {    }

} 
else 
{
    Compress-Archive -Path $folder\* -CompressionLevel Fastest -DestinationPath $logzipfile

    #deleting XML files
    Get-ChildItem $folder\*.xml | remove-item -Force
}

Write-Host "please collect 'C:\Temp\MSLogs\loggingFiles.zip' and send it to your support engineer." -ForegroundColor Cyan