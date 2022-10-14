<#
.NOTES
	Name: Export-MeetingItems.ps1
	Authors: Agustin Gallegos
   
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.SYNOPSIS
    Exports calendar items, some user/room mailboxes have in Exchange Online.

.DESCRIPTION
    Exports calendar items, some user/room mailboxes have in Exchange Online.
    Report will be exported to a ExportFolderPath by default to user's Desktop.

.PARAMETER Mailboxes
    List of SMTP email addresses to work with.

.PARAMETER ClientID
    This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

.PARAMETER TenantID
    This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

.PARAMETER ClientSecret
    This is an optional parameter. String parameter with the Client Secret which is configured in the AzureAD App.

.PARAMETER ExportFolderPath
    Insert target folder path named like "C:\Temp". By default this will be "$home\desktop"

.PARAMETER StartDate
    Set the start date to look for items. By default will consider 1 year backwards from the current date.

.PARAMETER EndDate
    Set the end date to look for items. By default will consider 1 year forwards from the current date.

.PARAMETER EnableTranscript
    Enable this parameter to write a powershell transcript in your 'Documents' folder.

.EXAMPLE
    PS C:\> .\Export-MeetingItems.ps1 -Mailboxes "user1@contoso.com" -EnableTranscript

    The script will ask for a user credential with impersonation permissions granted.
    will run against the "user1@contoso.com" mailbox and archive (if exists).
    Will Export meeting items to the default folder "$home\desktop" in a file named by 'user1-CalendaritemsReport.csv"


.EXAMPLE
    PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
    PS C:\> .\Export-MeetingItems.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -ExportFolderPath "C:\Reports" -EnableTranscript

    The script will collect all user's primary SMTP addresses from mailboxes belonging to "Staff" department (this command line would need to be connected to EXO Powershell).
    Will run against each mailbox and archive (if exists).
    Will Export meeting items to the selected folder "C:\Reports" in a file named by '<alias>-CalendaritemsReport.csv" for each user account.


.EXAMPLE
    PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "HR"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
    PS C:\> .\Export-MeetingItems.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -ExportFolderPath "C:\Reports" -ClientID "12345678" -TenantID "abcdefg" -ClientSecret "a1b2c3d4!#$"

    The script will collect all user's primary SMTP addresses from mailboxes belonging to "HR" department (this command line would need to be connected to EXO Powershell).
    Will connect using App-Only permission (this requires an AzureAD app and does not require Exchange Impersonation permissions. More info at: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
    Will run against each mailbox and archive (if exists).
    Will Export meeting items to the selected folder "C:\Reports" in a file named by '<alias>-CalendaritemsReport.csv" for each user account.


.COMPONENT
    STORE, Calendar

.ROLE
    Support
#>
param(
    [String[]] $mailboxes,

    [Parameter(Mandatory = $false, HelpMessage = 'Insert target folder path name like "C:\Temp"')]
    $ExportFolderPath = "$home\Desktop\",

    [DateTime] $StartDate = (Get-date).AddYears(-1),
    
    [DateTime] $EndDate = (Get-date).AddDays(364),

    [string] $ClientId,

    [String] $TenantID,

    [String] $ClientSecret,

    [switch]$EnableTranscript = $false
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

if ($EnableTranscript) {
    Start-Transcript
}

# creating folder path if it doesn't exists
if ( $ExportFolderPath -ne "$home\Desktop\" ) {
    if ( -not (Test-Path $ExportFolderPath) ) {
        write-host "Folder '$ExportFolderPath' does not exists. Creating folder." -foregroundColor Green
        $null = New-Item -Path $ExportFolderPath -ItemType Directory -Force
    }
}
else {
    # Checking if Desktop folder is located in the user's profile folder, or synched to OneDrive
    if ( -not(Test-Path $ExportFolderPath) ) {
        $ExportFolderPath = "$env:OneDriveCommercial\Desktop"
    }
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
    $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg", "")
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
if ( -not(Get-Module MSAL.PS -ListAvailable) ) {
    Install-Module MSAL.PS -Force -ErrorAction Stop
}
Import-Module MSAL.PS

# Connecting using Oauth with delegated permissions
if ( $clientID -eq '' -or $TenantID -eq '' -or $CertificateThumbprint -eq '' ) {                
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
}
# Connecting using Oauth with Application permissions
else {
    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    try {
        $global:token = Get-MsalToken -ClientId $clientID -TenantId $TenantID -ClientSecret $ClientSecret -Scopes $scopes -ErrorAction Stop
    }
    catch {
        if ( $_.Exception.Message -match "8856f961-340a-11d0-a96b-00c04fd705a2") {
            Write-Host "Known issue occurred. There is work in progress to fix authentication flow." -ForegroundColor red
            Write-Host "Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -ForegroundColor Red
            exit
        }
    }
}
$exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.AccessToken)
$service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
$Service.Credentials = $exchangeCredentials
$service.ReturnClientRequestId = $true
$service.UserAgent = "ExportMeetingItems/1.03"
#endregion

if ( $null -eq $mailboxes ) {
    $mailboxes = $token.Account.Username
}
$i=0
foreach ($mb in $mailboxes) {
    $i++
    Write-Progress -Id 0 -Activity "Processing mailbox: $i out of $($mailboxes.Count)" -status "Percent scanned: " -PercentComplete ($i * 100 / $($mailboxes.Count)) -ErrorAction SilentlyContinue

    $TargetSmtpAddress = $mb
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
    $service.HttpHeaders.Clear()
    $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)

    $filePath = "$ExportFolderPath\$($TargetSmtpAddress.split("@")[0])-CalendaritemsReport_$(Get-Date -Format "yyyy_MM_dd HH-mm-ss").csv"
    
    # binding to calendar in the primary mailbox, and archive mailbox if exists
    [int]$NumOfItems = 1000000
    $foldersToProcess = @()
    $Calendarfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
    $foldersToProcess += $Calendarfolder
    try {
        $ArchiveRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot)
        $archiveCalendarFolder = $archiveRoot.FindFolders($NumOfItems) | where-object Displayname -eq "Calendar"
        $foldersToProcess += $archiveCalendarFolder
    }
    catch {

    }

    # variables
    $calView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, $NumOfItems)
    $calView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
    $loopCount = -1

    foreach ($folder in $foldersToProcess) {
        $loopCount++
        $item = 0
        Write-Progress -Id 1 -ParentId 0 -activity "Processing folders: $($loopCount+1) out of $($foldersToProcess.Count)" -status "Percent scanned: " -PercentComplete ($($loopCount + 1) * 100 / $($foldersToProcess.Count)) -ErrorAction SilentlyContinue
        if ($loopCount -eq 0 ) { $mailboxProcessed = "$($TargetSmtpAddress.split("@")[0])-PrimaryMailbox" }
        elseif ($loopCount -eq 1 ) { $mailboxProcessed = "$($TargetSmtpAddress.split("@")[0])-ArchiveMailbox" }

        $Appointments = $folder.FindAppointments($calView)
        foreach ($Appointment in $Appointments.Items) {
            $item++
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing item: $item out of $($Appointments.Items.Count)" -status "Percent scanned: " -PercentComplete ($item * 100 / $($Appointments.Items.Count)) -ErrorAction SilentlyContinue
        
            $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $Appointment.Id)
            $Subject = $tempItem.subject.ToString().replace($tempItem.Organizer.Name, '')
            $output = $tempItem | Select-Object @{N="Mailbox";E={$mailboxProcessed}},@{N = "Subject"; E = { $Subject.trimstart() } }, organizer, @{N="RequiredAttendees";E={$_.RequiredAttendees -join ";"}}, @{N="OptionalAttendees";E={$_.OptionalAttendees -join ";"}}, @{N="Resources";E={$_.Resources -join ";"}}, start, end, isRecurring, appointmenttype, id
            $output | export-csv $filePath -NoTypeInformation -Append
        }
    }
}
if ($EnableTranscript) { stop-transcript }