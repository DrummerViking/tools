<#
.NOTES
	Name: Get-MRMRoamingXMLData.ps1
	Authors: Agustin Gallegos
    
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

    .SYNOPSIS
    Get MRM setting stamped in a user's mailbox.

    .DESCRIPTION
    Get MRM setting stamped in a user's mailbox.
    Output will be parsed from XML to readable text.

    .PARAMETER TargetSMTPAddress
    Use this optional parameter, to open a different mailbox. 
    You need to be assign Impersonation permissions, or FullAccess permisions in order to open another user's mailbox.

    .PARAMETER DeleteConfigurationMessage
    Using this parameter, deletes the IPM.Configuration.MRM message from the user mailbox.
    An Administrator should run 'Start-ManagedFolderAssistant' to issue MRM service and recreate the message.

    .EXAMPLE 
    PS C:\> Get-MRMRoamingXMLData.ps1
    In this example the script will ask for the user's credentials to be checked and get the MRM Roaming XML Data.

    .EXAMPLE 
    PS C:\> Get-MRMRoamingXMLData.ps1 -DeleteConfigurationMessage
    In this example the script will delete the 'IPM.Configuration.MRM' message from the user's mailbox.
    An Administrator should run Start-ManagedFolderAssistant to issue MRM service and recreate the message.

    .EXAMPLE
    PS C:\> Get-MRMRoamingXMLData.ps1 -TargetSMTPAddress 'anotherUser@domain.com'
    In this example the script will ask for the Admin's credentials to authenticate. And will actually open 'anotherUser@domain.com' mailbox to check and get the MRM Roaming XML Data.

    .COMPONENT
    STORE, MRM
    
    .ROLE
    Support
#>
[CmdletBinding()]
param(
    [String]$TargetSMTPAddress,

    [Switch]
    $DeleteConfigurationMessage
)
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
write-host " " 
Write-Host "This script requires at least EWS API 2.1" -ForegroundColor Yellow 

#region import EWS DLL file# Locating DLL location either in working path, in EWS API 2.1 path or in EWS API 2.2 path
$EWS = "$pwd\Microsoft.Exchange.WebServices.dll"
$test = Test-Path -Path $EWS
if ($test -eq $False) {
    Write-Host "EWS DLL in local path not found" -ForegroundColor Cyan
    $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
    if ( $null -eq $ewspkg ) {
        Write-Host "Downloading EWS DLL Nuget package and installing it" -ForegroundColor Cyan
        $null = Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -Force
        $null = Install-Package Microsoft.Exchange.WebServices -requiredVersion 2.2.0 -Scope CurrentUser
        $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
    }        
    $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg","")
    Write-Host "EWS DLL found in package folder path" -ForegroundColor Cyan
    $EWS = "$EWSPath\lib\40\Microsoft.Exchange.WebServices.dll"
}
else {
    Write-Host "EWS DLL found in current folder path" -ForegroundColor Cyan
}
Add-Type -Path $EWS
#endregion

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 

#Getting oauth credentials
if ( -not(Get-Module MSAL.PS -ListAvailable) ) {
    Install-Module MSAL.PS -Force -ErrorAction Stop
}
Import-Module MSAL.PS

# Connecting using Oauth with delegated permissions              
$ClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
$RedirectUri = "http://localhost/code"
$scopes = New-Object System.Collections.Generic.List[string]
$scopes.Add("https://outlook.office365.com/.default")
#$scopes.Add("https://outlook.office.com/EWS.AccessAsUser.All")
try {
    $token = Get-MsalToken -ClientId $clientID -RedirectUri $RedirectUri -Scopes $scopes -Interactive -ErrorAction Stop
}
catch {
    if ( $_.Exception.Message -match "8856f961-340a-11d0-a96b-00c04fd705a2") {
        Write-Host "Known issue occurred. There is work in progress to fix authentication flow." -ForegroundColor red
        Write-Host "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -ForegroundColor Red
        exit
    }
}
$exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.AccessToken)
$service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
$Service.credentials = $exchangeCredentials
$service.ReturnClientRequestId = $true
$service.UserAgent = "GetMRMRoamingXMLData/2.30"

if ( $TargetSMTPAddress ){
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
    $service.HttpHeaders.Clear()
    $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
    $SmtpAddress = $TargetSmtpAddress
}

if ( $DeleteConfigurationMessage )
{
    try
    {
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$SmtpAddress)
        $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "MRM", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
        $UsrConfig.Delete()
        write-host "Successfully deleted the MRM Configuration message." -ForegroundColor Green
    }
    catch
    {
        write-host "Configuration message does not exists, or it failed to be deleted." -ForegroundColor Red
    }
}
else
{
    try {
        # following script block can read the PR_ROAMING_XMLSTREAM text and parse it in a readable way
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$SmtpAddress)
        $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "MRM", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
        $ConfXML = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)
        [XML]$ConfXML = $ConfXML
        write-host "Following are current 'Delete' TAGs" -ForegroundColor Green
        $ConfXML.UserConfiguration.Info.Data.PolicyTag
        write-host ""
        write-host "Following are current 'Move to Archive' TAGs" -ForegroundColor Green
        $ConfXML.UserConfiguration.Info.Data.ArchiveTag
        write-host ""
        write-host "Following are current 'Default Archive TAGs" -ForegroundColor Green
        $ConfXML.UserConfiguration.Info.Data.DefaultArchiveTag
    }
    catch {
        write-host "Configuration message does not exists, or it failed to be read." -ForegroundColor Red
    }
}