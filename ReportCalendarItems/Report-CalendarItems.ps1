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
$EWS = "$PsscriptRoot\Microsoft.Exchange.WebServices.dll"
$test = Test-Path -Path $EWS
if ($test -eq $False) {
    Write-Host "EWS DLL in local path not found" -ForegroundColor Cyan
    $test2 = Test-Path -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
    if ($test2 -eq $False) {
        Write-Host "EWS 2.1 not found" -ForegroundColor Cyan
        $test3 = Test-Path -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        if ($test3 -eq $False) {
            Write-Host "EWS 2.2 not found" -ForegroundColor Cyan
        }
        else {
            Write-Host "EWS 2.2 found" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "EWS 2.1 found" -ForegroundColor Cyan
    }        
}
else {
    Write-Host "EWS DLL found in local path" -ForegroundColor Cyan
}
    
    
if ($test -eq $False -and $test2 -eq $False -and $test3 -eq $False) {
    Write-Host " "
    Write-Host "You don't seem to have EWS API dll file 'Microsoft.Exchange.WebServices.dll' in the same Directory of this script" -ForegroundColor Red
    Write-Host "please get a copy of the file or download the whole API from: " -ForegroundColor Red -NoNewline
    Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=42951" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "we will open your browser in 10 seconds automatically directly to this URL" -ForegroundColor Red
    Start-Sleep 10 
    Start-Process -FilePath "https://www.microsoft.com/en-us/download/details.aspx?id=42951"

    return
}
    
Write-host "EWS API detected. All good!" -ForegroundColor Cyan
            
if ($test -eq $True) {
    Unblock-File -Path $EWS -Confirm:$false
    Add-Type -Path $EWS
    Write-Host "Using EWS DLL in local path" -ForegroundColor Cyan
}
elseif ($test2 -eq $True) {
    Add-Type -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
    Write-Host "Using EWS 2.1" -ForegroundColor Cyan
}
elseif ($test3 -eq $True) {
    Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
    Write-Host "Using EWS 2.2" -ForegroundColor Cyan
}
write-host " "
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