<#
    .SYNOPSIS
    Script to replace room locations in user's meeting, with a new room location.
    
    .DESCRIPTION
    Script to replace room locations in user's meeting, with a new room location.
    
    .PARAMETER RoomsCSVFilePath
    Sets the Rooms mapping file path. This file should have 2 columns named "PreviousRoom","newRoom".
    
    .PARAMETER MailboxesCSVFilePath
    Sets the users file path. This file should have 1 column named "PrimarySMTPAddress".
    
    .PARAMETER StartDate
    Sets the start date to look for meeting item in the user mailboxes. By default is the current date.
    
    .PARAMETER EndDate
    Sets the end date to look for meeting item in the user mailboxes. By default is 1 year after the current date.
    
    .PARAMETER ValidateUsersExistence
    If this Switch parameter is used, the script will not only connect using EWS, but will attempt to connect to EXO Powershell module and validate the user mailboxes exists as valid mailboxes in Exchange Online.

    .PARAMETER ValidateRoomsExistence
    If this Switch parameter is used, the script will not only connect using EWS, but will attempt to connect to EXO Powershell module and validate the room mailboxes exists as valid recipients in Exchange Online.
    
    .PARAMETER NoTranscript
    If this Switch parameter is used, No Powershell Transcript will be created. By Default all information displayed in the Powershell console, will be exported to the transcript file usually saved in "Documents" folder.
    
    .EXAMPLE
    PS C:\> .\Replace-RoomsInMeetings.ps1 -ValidateRoomExistence

    In this example the script will pop-up and prompt for the CSV with the mapping file for room accounts, and the CSV file for the users where to replace the rooms.
    Aside of connecting to EWS, the script will connect to EXO Powershell (it might ask for credentials again) and validate the rooms detailed in the mapping file exists as recipients in EXO.
    the script will look for meeting items since the current day and 1 year forward.
    A transcript file will be created in the user's Documents folder.

    .EXAMPLE
    PS C:\> .\Replace-RoomsInMeetings.ps1 -RoomsCSVFilePath C:\Temp\RoomsMappingFile.csv

    In this example the script reads the Rooms mapping file from "C:\Temp\RoomsMappingFile.csv".
    Then will pop-up and prompt for the CSV file for the users where to replace the rooms.
    the script will look for meeting items since the current day and 1 year forward.
    A transcript file will be created in the user's Documents folder.

    .EXAMPLE
    PS C:\> .\Replace-RoomsInMeetings.ps1 -RoomsCSVFilePath C:\Temp\RoomsMappingFile.csv -MailboxesCSVFilePath C:\Temp\Users.Csv -EndDate 01/01/2025

    In this example the script reads the Rooms mapping file from "C:\Temp\RoomsMappingFile.csv" and user's list from "C:\Temp\Users.Csv".
    the script will look for meeting items since the current day through January 1st 2025.
    A transcript file will be created in the user's Documents folder.
#>
[CmdletBinding()]
param(
    [String] $RoomsCSVFilePath,

    [String] $MailboxesCSVFilePath,

    [DateTime]$StartDate = (get-date).ToShortDateString(),

    [DateTime]$EndDate = (get-date).AddYears(1).ToShortDateString(),

    [switch] $ValidateUsersExistence,

    [switch] $ValidateRoomsExistence,

    [switch] $NoTranscript
)

Begin {
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
    if ($NoTranscript -eq $False) {
        Start-Transcript
    }
    $EWS = "$pwd\Microsoft.Exchange.WebServices.dll"
    $test = Test-Path -Path $EWS
    if ($test -eq $False) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL in local path not found" -ForegroundColor Cyan
        $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        if ( $null -eq $ewspkg ) {
            Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Downloading EWS DLL Nuget package and installing it" -ForegroundColor Cyan
            $null = Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -Force
            $null = Install-Package Microsoft.Exchange.WebServices -requiredVersion 2.2.0 -Scope CurrentUser
            $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        }        
        $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg","")
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in package folder path" -ForegroundColor Cyan
        $EWS = "$EWSPath\lib\40\Microsoft.Exchange.WebServices.dll"
    }
    else {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in current folder path" -ForegroundColor Cyan
    }
    Add-Type -Path $EWS
}

process {
    #region Selecting Rooms CSV file
    if ( $RoomsCSVFilePath.Length -eq 0 ) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Please pick up the CSV files with the list of previous and new rooms to replace."
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        [System.Windows.Forms.Application]::EnableVisualStyles() 
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $PSScriptRoot
        $OpenFileDialog.ShowDialog() | Out-Null
        if ($OpenFileDialog.filename -ne "") {
            $RoomsCSVPath = $OpenFileDialog.filename
        }
    }else {
        $RoomsCSVPath = $RoomsCSVFilePath
    }
    $csv = Import-csv $RoomsCSVPath
    if ( ($csv | Get-Member -Name PreviousRoom,NewRoom).count -lt 2 ){
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Rooms mailboxes CSV file does not contain the necessary columns. Please check you have 'PreviousRoom, NewRoom' columns and try again." -ForegroundColor Red
        return
    }
    [pscustomobject]$rooms = @{}
    foreach ( $room in $csv ) {
        $rooms[$room.PreviousRoom] = $room.newRoom
    }
    #endregion

    #region Importing mailboxes list
    if ( $MailboxesCSVFilePath.Length -eq 0 ) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Please pick up the CSV files with the list of mailboxes to search for meetings to be updated."
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        [System.Windows.Forms.Application]::EnableVisualStyles() 
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $PSScriptRoot
        $OpenFileDialog.ShowDialog() | Out-Null
        if ($OpenFileDialog.filename -ne "") {
            $MailboxesCSV = $OpenFileDialog.filename
        }
    }else {
        $MailboxesCSV = $MailboxesCSVFilePath
    }
    $mbxs = Import-Csv $MailboxesCSV
    if ( ($mbxs | Get-Member -Name PrimarySMTPAddress).count -lt 1 ){
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Mailboxes CSV file does not contain the necessary column. Please check you have 'PrimarySMTPAddress' column and try again." -ForegroundColor Red
        return
    }
    #endregion

    #creating service object
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
    
    #region Getting oauth credentials using MSAL
    Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to EWS. Please insert user credentials with Impersonation permissions."
    if ( -not(Get-Module Microsoft.Identity.Client -ListAvailable) ) {
        Install-Module Microsoft.Identity.Client -Force -ErrorAction Stop
    }
    Import-Module Microsoft.Identity.Client -Force -ErrorAction SilentlyContinue
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
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Known issue occurred. There is work in progress to fix authentication flow." -ForegroundColor red
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Failed to obtain authentication token. Exiting script. Please rerun the script again and it should work." -ForegroundColor Red
        exit
    }
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.Result.AccessToken)
    $service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
    $Service.Credentials = $exchangeCredentials
    #endregion

    #region Validate if user mailboxes exists as valid recipients in EXO
    if ( $ValidateUsersExistence ) {
        if ( (Get-PSSession).Computername -notcontains "outlook.office365.com" ) {
            if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable) ) {
                Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement
            Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online. Please insert credentials with Exchange Admin Role."
            Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
        }
        foreach ($user in $mbxs) {
            try {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Checking if user mailbox $($user.PrimarySMTPAddress) exists..." -NoNewline
                $null = Get-EXOMailbox $user.PrimarySMTPAddress -ErrorAction Stop
                Write-host "Ok." -ForegroundColor Green
            }
            catch {
                Write-host "Failed." -ForegroundColor Red
                $failedAlias = $_.Exception.Message.split("'")[2]
                Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] User mailbox '$failedAlias' not found. Exiting script." -ForegroundColor Red
                exit
            }
        }      
    }
    #endregion
    
    #region Validate if room mailboxes exists as valid recipients in EXO
    if ( $ValidateRoomsExistence ) {
        if ( (Get-PSSession).Computername -notcontains "outlook.office365.com" ) {
            if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable) ) {
                Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement
            Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online. Please insert credentials with Exchange Admin Role."
            Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
        }
        foreach ($line in $csv) {
            try {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Checking if Room mailbox $($line.NewRoom) exists..." -NoNewline
                $null = Get-EXORecipient $line.newRoom -ErrorAction Stop
                Write-host "Ok." -ForegroundColor Green

                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Checking if Room mailbox $($line.PreviousRoom) exists..." -NoNewline
                $null = Get-EXORecipient $line.PreviousRoom -ErrorAction Stop
                Write-host "Ok." -ForegroundColor Green
            }
            catch {
                Write-host "Failed." -ForegroundColor Red
                $failedAlias = $_.Exception.Message.split("'")[2]
                Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Room mailbox '$failedAlias' not found. Exiting script." -ForegroundColor Red
                exit
            }
        }      
    }
    #endregion

    # loop through each mailbox user looking for meeting items
    $i = 0
    $mbxsCount = 1
    if ( $mbxs.Count -gt 0 ) {
        $mbxsCount = $mbxs.count
    }
    foreach ($mbx in $mbxs) {
        Remove-Variable Appointments -Force -ErrorAction SilentlyContinue
        $i++
        $j = 0
        Write-Progress -Id 0 -Activity "Scanning mailbox $i out of $mbxsCount" -status "Percent scanned: " -PercentComplete ($i * 100 / $mbxsCount) -ErrorAction SilentlyContinue
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Working on mailbox: $($mbx.PrimarySMTPAddress)" -ForegroundColor Green
        # Setting impersonation address to target mailbox
        $TargetSmtpAddress = $mbx.PrimarySMTPAddress
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Clear()
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
    
        try {
            $Calendarfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        }
        catch {
            Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Something failed to connect to the mailbox: $($mbx.PrimarySMTPAddress)" -ForegroundColor Red
            continue
        }
        
        [int]$NumOfItems = 10000

        $calView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, $NumOfItems)
        $calView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)

        # loop through each calendar item in the current mailbox
        $Appointments = $Calendarfolder.FindAppointments($calView)
        foreach ($Appointment in $Appointments) {
            $j++
            Write-Progress -Id 0 -Activity "Scanning item $j out of $($Appointments.Items.count)" -status "Percent scanned: " -PercentComplete ($j * 100 / $($Appointments.Items.count)) -ErrorAction SilentlyContinue
            try {
                $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $Appointment.Id)
                Write-Verbose "[$((Get-Date).ToString("HH:mm:ss"))] Scanning item: '$($tempItem.Subject)'"
                $roomFound = $csv.previousRoom -eq $tempItem.Resources.Address
                if ( $roomFound.GetType().name -eq "Boolean" -and $True -eq $roomFound ) {
                    $roomFound = $csv.previousRoom
                }
                elseif ($roomFound.GetType().name -eq "Boolean" -and $False -eq $roomFound) {
                    $roomFound = @()
                }
                # If resources is empty
                # OR If resources is not empty but does not contain any of the PreviousRoom accounts we want to replace
                # OR if the user being scanned is not the current Organizer
                #   THEN we will continue to the next calendar item
                if ( $tempItem.Resources.Count -eq 0 -or $roomFound.count -eq 0 -or $tempItem.Organizer.Address -ne $TargetSmtpAddress) {
                    if ($tempItem.Resources.Count -eq 0) {
                        Write-Verbose "[$((Get-Date).ToString("HH:mm:ss"))] Skipping item '$($tempItem.Subject)' because it doesn't have resources, looks to be a single appointment."
                    }
                    elseif ($roomFound.count -eq 0) {
                        Write-Verbose "[$((Get-Date).ToString("HH:mm:ss"))] Skipping meeting '$($tempItem.Subject)' because it doesn't have any resources listed in the rooms mapping CSV file."
                    }
                    elseif ($tempItem.Organizer.Address -ne $TargetSmtpAddress) {
                        Write-Verbose "[$((Get-Date).ToString("HH:mm:ss"))] Skipping meeting '$($tempItem.Subject)' because current working mailbox is not the organizer and cannot update the meeting."
                    }
                    continue
                }
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Previous room $roomFound found in meeting '$($tempItem.Subject)'." -ForegroundColor Cyan
                # doing a Check Name against the new Room mailbox
                $recipientResolved = $service.ResolveName($rooms[$roomFound])
                $newRoomAttendee = New-Object Microsoft.Exchange.WebServices.Data.Attendee($recipientResolved.mailbox.Address)
                $newRoomAttendee.RoutingType = $recipientResolved.mailbox.RoutingType
                $newRoomAttendee.Name = $recipientResolved.mailbox.Name
                
                # if the calendar item is a recurring meeting, we are binding to the recurrening Master item, in order to successfully update the whole series.
                if ( $tempItem.IsRecurring -eq $true ) {
                    Write-Verbose "[$((Get-Date).ToString("HH:mm:ss"))] Processing item as a recurrent meeting '$($tempItem.Subject)'. Updating Recurring Master item."
                    $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($service, $tempItem.Id)
                }
                $tempItem.Resources.Clear()
                $null = $tempItem.Resources.Add($newRoomAttendee)
                $tempItem.Location = $newRoomAttendee.Name
                $tempItem.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy)
                write-host "[$((Get-Date).ToString("HH:mm:ss"))] Replacing $roomFound for $($rooms[$roomFound])" -ForegroundColor Cyan
            }
            catch {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Something went wrong to update meeting '$($Appointment.Subject)' on mailbox $TargetSmtpAddress. Error message: $_"
                continue
            }
        }
    }
}
End {
    if ($NoTranscript -eq $False) {
        Stop-Transcript
    }
}