#Requires -RunAsAdministrator
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
[CmdletBinding()]
Param(
    [Parameter(Position=1,Mandatory = $True, HelpMessage = 'Primary Email Address of an on-premises mailbox you want to check...')]
    [string]$OnpremisesUser = '',

    [Parameter(Position=2,Mandatory = $True, HelpMessage = 'Primary Email Address of a cloud mailbox you want to check...')]
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

# exporting current Auth Certificate, in order to check MsolServicePrincipalCredential later in MSOnline
Get-AuthConfig | export-clixml -path "$ts.OnPrem_Authconfig.xml"
$thumbprint = (Get-AuthConfig).CurrentCertificateThumbprint
$oAuthCert = (Get-ChildItem Cert:\LocalMachine\My) | Where-Object {$_.Thumbprint -match $thumbprint}
$certType = [System.Security.Cryptography.X509Certificates.X509ContentType]::Cert
$certBytes = $oAuthCert.Export($certType)
$CertFile = "C:\temp\MSlogs\OAuthCert.cer"
[System.IO.File]::WriteAllBytes($CertFile, $certBytes)

# exporting Outputs
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting AuthServer info"
Get-AuthServer | export-clixml -path "$ts.OnPrem_AuthServer.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting PartnerApplication"
Get-PartnerApplication | export-clixml -path "$ts.OnPrem_PartnerApplication.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting EWS Virtual Directories"
Get-WebServicesVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_EWSVDir.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Autodiscover Virtual Directories"
Get-AutoDiscoverVirtualDirectory -ShowMailboxVirtualDirectories | export-clixml -path "$ts.OnPrem_AutodVdir.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting IntraOrganizationConnectors"
Get-IntraOrganizationConnector | export-clixml -path "$ts.OnPrem_IOC.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Availability Address Spaces"
Get-AvailabilityAddressSpace | export-clixml -path "$ts.OnPrem_AvailabilityAddressSpaces.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Remote Mailbox info"
Get-RemoteMailbox $CloudUser | export-clixml -path "$ts.OnPrem_RemoteMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting On-premises Mailbox info"
Get-Mailbox $OnpremisesUser | export-clixml -path "$ts.OnPrem_OnPremisesMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Testing OAUTH Connectivity to EWS service"
Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $OnpremisesUser -Verbose | export-clixml -path "$ts.OnPrem_TestOAuthEWS.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting OAUTH Connectivity to AutoD service"
Test-OAuthConnectivity -Service AutoD  -TargetUri https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Mailbox $OnpremisesUser -Verbose | export-clixml -path "$ts.OnPrem_TestOAuthAutoD.xml"

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

#checking if MSOnline module is installed
#    if not installed, and if we are running PSVersion5, we will attempt to install it
#    if not installed, and we are running PS3 or PS4, we check if we have 'PowershellGet' installed
#        if PSGet is not installed, we will request to download and install it.
#    if not installed, and we are running PS3 or PS4 and now we have PSGet installed, we will attempt to install MSOnline module
#at the end we will attempt to import and connect to MSOnline    

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to MSOnline" 
if($null -eq (get-module -ListAvailable -Name msonline)){
    if($PSVersionTable.PSVersion.Major -ge 5){
        Install-Module msonline -Force -Confirm:$False
    }else{
        if($null -eq (get-module -ListAvailable -Name PowershellGet)){
            write-host "MSONLINE module has not been detected. Please install 'https://www.microsoft.com/en-us/download/details.aspx?id=51451' and re run this script" -ForegroundColor White -BackgroundColor Red
            write-host "Installing PowershellGet will allow this script to download and install MSONLINE module in this machine." -ForegroundColor White -BackgroundColor Red
            Start-Process https://www.microsoft.com/en-us/download/details.aspx?id=51451
            return
        }else{
            Install-Module msonline -Force -Confirm:$False
        }
    }
}
Import-Module msonline
Connect-MsolService -Credential $LiveCred

$FormatEnumerationLimit = -1

# getting AuthCertificate key value
$objFSO = New-Object -ComObject Scripting.FileSystemObject;
$CertFile = $objFSO.GetAbsolutePathName($CertFile);
$cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate
$cer.Import($CertFile);
$binCert = $cer.GetRawCertData();
$credValue = [System.Convert]::ToBase64String($binCert);
$ServiceName = "00000002-0000-0ff1-ce00-000000000000";
$p = Get-MsolServicePrincipal -ServicePrincipalName $ServiceName

# exporting Outputs
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting MSOL Service Principal Credentials"
Get-MsolServicePrincipalCredential -ObjectId $p.ObjectId -ReturnKeyValues $true | Where-Object{$_.Value -eq $credValue} | export-clixml -path "$ts.Cloud_MsolServicePrincipalCredential.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting MSOL Service Principals"
$p | export-clixml -path "$ts.Cloud_MsolServicePrincipal.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's IntraOrganizationConnector"
Get-EOIntraOrganizationConnector | export-clixml -path "$ts.Cloud_IOC.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Mail User info"
Get-EOMailUser $OnpremisesUser | export-clixml -path "$ts.Cloud_OnPremisesMBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Mailbox info"
Get-EOMailbox $CloudUser | export-clixml -path "$ts.Cloud_MBX.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Inbound Connector"
Get-EOInboundConnector | export-clixml -path "$ts.Cloud_InboundConnectors.xml"

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting cloud's Outbound Connector"
Get-EOOutboundConnector | export-clixml -path "$ts.Cloud_OutboundConnectors.xml"

Write-Host ""
Write-Host "if you want to test OAUTH, please enter your on-premises EWS and AutoD published FQDN. For Example 'mail.contoso.com'." -ForegroundColor Yellow
Write-Host "if you want to skip the test, leave blank and hit Enter key: " -NoNewline -ForegroundColor Yellow
$ewsdomain = Read-Host -Prompt "Enter EWS endpoint"
if($ewsdomain -ne "")
{
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Testing OAUTH Connectivity to on-premises EWS service"
    Test-EOOAuthConnectivity -Service EWS -TargetUri https://$ewsdomain/ews/exchange.asmx -Mailbox $CloudUser -Verbose | export-clixml -path "$ts.Cloud_TestOAuthEWS.xml"
}
$autodomain = Read-Host -Prompt "Enter Autodiscover endpoint. for example autodiscover.contoso.com"
if($autodomain -ne "")
{
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Testing OAUTH Connectivity to on-premises AutoD service"
    Test-EOOAuthConnectivity -Service AutoD -TargetUri https://$autodomain/autodiscover/autodiscover.svc -Mailbox $CloudUser -Verbose | export-clixml -path "$ts.Cloud_TestOAuthAutoD.xml"
}

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