<#
.NOTES
	Name: Report-CalendarItems.ps1
	Authors: Agustin Gallegos
   
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
    Report how many calendar items, per calendar year, some user/room mailboxes have in Exchange Online.
.DESCRIPTION
    Reports how many calendar items, per calendar year, some user/room mailboxes have in Exchange Online.
    Report can be exported to a DestinationFolderPath or by default to user's Desktop.
.PARAMETER EnableTranscript
    Enable this parameter to write a powershell transcript in your 'Documents' folder.
.PARAMETER CSVFile
    CSV file must contain a unique header named "PrimarySMTPAddress".
.PARAMETER DestinationFolderPath
    Insert target folder path name like "C:\Temp".
.EXAMPLE 
    .\Report-CalendarItems.ps1
.EXAMPLE 
    .\Report-CalendarItems.ps1 -EnableTranscript
.EXAMPLE    
    .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -EnableTranscript
.EXAMPLE    
    .\Report-CalendarItems.ps1 -CSVFile "D:\Temp\rooms.csv" -DestinationFolderPath "C:\Reports" -EnableTranscript

.COMPONENT
   STORE, Calendar
.ROLE
   Support
#>
param(
    [switch]$EnableTranscript = $false,

    [Parameter(Mandatory = $false, HelpMessage = 'CSV file must contain a unique header named "PrimarySMTPAddress"')]$CSVFile = $Null,
    
    [Parameter(Mandatory = $false, HelpMessage = 'Insert target folder path name like "C:\Temp"')]$DestinationFolderPath = "$home\Desktop"
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
# profits, business interruption, loss of business information, or other pecuniary loss 
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
# 
#################################################################################
"@
Write-Host $disclaimer -foregroundColor Yellow
Write-Host " " 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#$psCred = Get-Credential -Message "Type your Service account's credentials"

if ($EnableTranscript) {
    Start-Transcript
}

#region load EWS API DLL
write-host " " 
Write-Host "This script requires at least EWS API 2.1" -ForegroundColor Yellow 
 
# Locating DLL location either in working path, in EWS API 2.1 path or in EWS API 2.2 path
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

#region Create Service Object
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
#Getting oauth credentials
if ( !(Get-Module Microsoft.Identity.Client -ListAvailable) -and !(Get-Module Microsoft.Identity.Client) ) {
    Install-Module Microsoft.Identity.Client -Force -ErrorAction Stop
}
Import-Module Microsoft.Identity.Client
$AppId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
$pcaOptions = [Microsoft.Identity.Client.PublicClientApplicationOptions]::new()
$pcaOptions.ClientId = $AppId
$pcaOptions.RedirectUri = "http://localhost/code"
$pcaBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($pcaOptions)
$pca = $pcaBuilder.Build()
$scopes = New-Object System.Collections.Generic.List[string]
$scopes.Add("https://outlook.office365.com/.default")
#$scopes.Add("https://outlook.office.com/EWS.AccessAsUser.All")
$authResult = $pca.AcquireTokenInteractive($scopes)
$token = $authResult.ExecuteAsync()
while ( $token.IsCompleted -eq $False ) { <# Waiting for token auth flow to complete #>}
if ($token.Status -eq "Faulted" -and $token.Exception.Message.StartsWith("One or more errors occurred. (ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2'")) {
    Write-Host "Known issue occurred. There is work in progress to fix authentication flow." -ForegroundColor red
    Write-Host "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -ForegroundColor Red
    exit
}
$exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.Result.AccessToken)
$service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
$Service.Credentials = $exchangeCredentials
#endregion

# Selecting CSV file if it is not pass as a Parameter
if ( $null -eq $CSVFile ) {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.ShowDialog() | Out-Null
    if ($OpenFileDialog.filename -ne "") {
        $CSVFile = $OpenFileDialog.filename
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Select file Operation finished" -ForegroundColor Yellow
    }
}

# Importing CSV file to powershell
$mbxs = Import-Csv -Path $CSVFile
$mailboxCount = $mbxs.Count
$i = 0
foreach ($mbx in $mbxs) {
    $i++
    Write-Progress -activity "Scanning Users: $i out of $mailboxCount" -status "Percent scanned: " -PercentComplete ($i / $mailboxCount * 100) -ErrorAction SilentlyContinue

    $TargetSmtpAddress = $mbx.PrimarySMTPAddress
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
    $service.HttpHeaders.Clear()
    $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        
    $Calendarfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
    
    # looping through the last 5 years
    [int]$currentYear = (get-date).Year
    [int]$startYear = (Get-Date).AddYears(-5).Year
    $currentYear..$startYear | ForEach-Object {
        $startDate = "01/01/$_"
        $endDate = "12/31/$_"
        [int]$NumOfItems = 10000
        
        $calView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, $NumOfItems)
        $calView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)

        $Appointments = $Calendarfolder.FindAppointments($calView)
        foreach ($Appointment in $Appointments) {
            $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $Appointment.Id)
            $Subject = $tempItem.subject.ToString().replace($tempItem.Organizer.Name, '')
            $output = $tempItem | Select-Object @{N = "Mailbox"; E = { $tempItem.LastModifiedName } }, @{N = "Subject"; E = { $Subject.trimstart() } }, organizer, start, end, datetimereceived
            $output | export-csv "$DestinationFolderPath\YearList-$_.csv" -NoTypeInformation -Append
        }
    }
}
if ($EnableTranscript) { stop-transcript }