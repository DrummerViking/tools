<#
.NOTES
	Name: DeleteMeetings-GUI.ps1
	Authors: Agustin Gallegos
	Version History:
	1.00 - 12/27/2018 - Project start
    1.00 - 12/27/2018 - First Release
    1.30 - 01/03/2019 - Remove hardcoded timeframe of 180 days. Now user can select desired time frame, including past items.
                        Added 'Subject' column to results.**Take into account, EXO's default Calendar Processing is to Delete the Subject for Room Mailboxes**
    
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
.SYNOPSIS
    Delete Meetings items for Organizers that already left the company, in Exchange Online.
.DESCRIPTION
    Delete Meetings items for Organizers that already left the company, in Exchange Online.
    You can pass a list of users/room mailboxes, and delete all meetings found from a specific Organizer. 
.EXAMPLE 
    .\DeleteMeetings-GUI.ps1
    .\DeleteMeetings-GUI.ps1 -EnableTranscript

.COMPONENT
   STORE, Calendar
.ROLE
   Support
#>
param(
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

#$psCred = Get-Credential -Message "Type your Service account's credentials"
function GenerateForm {
 
#Internal function to request inputs using UI instead of Read-Host
function Show-InputBox{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $Prompt,
        
        [Parameter(Mandatory=$false)]
        [string]
        $DefaultValue='',
        
        [Parameter(Mandatory=$false)]
        [string]
        $Title = 'Windows PowerShell'
    )
    
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.Interaction]::InputBox($Prompt,$Title, $DefaultValue)
}
 
#region Import the Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles() 
#endregion

#region Generated Form Objects
$MainForm = New-Object System.Windows.Forms.Form
$radiobutton1 = New-Object System.Windows.Forms.RadioButton
$txtBoxMbxAlias = New-Object System.Windows.Forms.TextBox
$radiobutton2 = New-Object System.Windows.Forms.RadioButton
$buttonImportFile = New-Object System.Windows.Forms.Button
$labImportFileHowTo = New-Object System.Windows.Forms.Label

$labFromDate = New-Object System.Windows.Forms.Label
$FromDatePicker = New-Object System.Windows.Forms.DateTimePicker
$labToDate = New-Object System.Windows.Forms.Label
$ToDatePicker = New-Object System.Windows.Forms.DateTimePicker
$labelOrganizer = New-Object System.Windows.Forms.Label
$txtBoxOrganizer = New-Object System.Windows.Forms.TextBox
$buttonList = New-Object System.Windows.Forms.Button
$buttonDelete = New-Object System.Windows.Forms.Button

$buttonExit = New-Object System.Windows.Forms.Button
$labelHelp = New-Object System.Windows.Forms.Label
$dgResults = New-Object System.Windows.Forms.DataGridView 
$txtBoxResults = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects
 
if($EnableTranscript){
    Start-Transcript
}

#region load EWS API DLL
write-host " " 
Write-Host "This script requires at least EWS API 2.1" -ForegroundColor Yellow 
 
    # Locating DLL location either in working path, in EWS API 2.1 path or in EWS API 2.2 path
    $Directory = ".\"
    $EWS = Join-Path $Directory "Microsoft.Exchange.WebServices.dll"
    $test = Test-Path -Path $EWS
    if ($test -eq $False){
        Write-Host "EWS DLL in local path not found" -ForegroundColor Cyan
        $test2 = Test-Path -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        if ($test2 -eq $False){
            Write-Host "EWS 2.1 not found" -ForegroundColor Cyan
            $test3 = Test-Path -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
            if ($test3 -eq $False) {
                Write-Host "EWS 2.2 not found" -ForegroundColor Cyan
                }else{
                Write-Host "EWS 2.2 found" -ForegroundColor Cyan
                }
              }else{
            Write-Host "EWS 2.1 found" -ForegroundColor Cyan
            }        
        }else{
        Write-Host "EWS DLL found in local path" -ForegroundColor Cyan
        }
    
    
    if($test -eq $False -and $test2 -eq $False -and $test3 -eq $False){
             Write-Host " "
          Write-Host "You don't seem to have EWS API dll file 'Microsoft.Exchange.WebServices.dll' in the same Directory of this script" -ForegroundColor Red
          Write-Host "please get a copy of the file or download the whole API from: " -ForegroundColor Red -NoNewline
          Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=42951" -ForegroundColor Cyan
          Write-Host ""
          Write-Host "we will open your browser in 10 seconds automatically directly to this URL" -ForegroundColor Red
          sleep 10 
          Start-Process -FilePath "https://www.microsoft.com/en-us/download/details.aspx?id=42951"

          return
    }
    
    Write-host "EWS API detected. All good!" -ForegroundColor Cyan
            
    if ($test -eq $True){
        Add-Type -Path $EWS
        Write-Host "Using EWS DLL in local path" -ForegroundColor Cyan
        }
    elseif($test2 -eq $True){
        Add-Type -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        Write-Host "Using EWS 2.1" -ForegroundColor Cyan
        }
    elseif ($test3 -eq $True){
        Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.*\Microsoft.Exchange.WebServices.dll"
        Write-Host "Using EWS 2.2" -ForegroundColor Cyan
        }
    write-host " "
#endregion


#region Select Exchange version and establish connection

    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm = New-Object System.Windows.Forms.Form
    $radiobutton3 = New-Object System.Windows.Forms.RadioButton
    $radiobutton4 = New-Object System.Windows.Forms.RadioButton
    $buttonGo = New-Object System.Windows.Forms.Button

    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($radiobutton3)
    $PremiseForm.Controls.Add($radiobutton4)
    $PremiseForm.ClientSize = New-Object System.Drawing.Size(250,160)
    $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $PremiseForm.Name = "form1"
    $PremiseForm.Text = "Choose your Exchange version"
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton1.Location = New-Object System.Drawing.Point(20,20)
    $radiobutton1.Size = New-Object System.Drawing.Size(150,25)
    $radiobutton1.TabStop = $True
    $radiobutton1.Text = "Exchange 2007"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20,50)
    $radiobutton2.Size = New-Object System.Drawing.Size(150,20)
    $radiobutton2.TabStop = $True
    $radiobutton2.Text = "Exchange 2010"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True
    #
    # radiobutton3
    #
    $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton3.Location = New-Object System.Drawing.Point(20,80)
    $radiobutton3.Size = New-Object System.Drawing.Size(150,25)
    $radiobutton3.TabStop = $True
    $radiobutton3.Text = "Exchange 2013/2016"
    $radiobutton3.Checked = $false
    $radiobutton3.UseVisualStyleBackColor = $True
    #
    # radiobutton4
    #
    $radiobutton4.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton4.Location = New-Object System.Drawing.Point(20,110)
    $radiobutton4.Size = New-Object System.Drawing.Size(150,30)
    $radiobutton4.Text = "Office365"
    $radiobutton4.Checked = $false
    $radiobutton4.UseVisualStyleBackColor = $True

    #"Go" button
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 170
    $System_Drawing_Point.Y = 20
    $buttonGo.Location = $System_Drawing_Point
    $buttonGo.Name = "Go"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 25
    $System_Drawing_Size.Width = 50
    $buttonGo.Size = $System_Drawing_Size
    $buttonGo.Text = "Go"
    $buttonGo.UseVisualStyleBackColor = $True
    $buttonGo.add_Click({
        if($radiobutton1.Checked){$Global:option = "Exchange2007_SP1"}
        elseif($radiobutton2.Checked){$Global:option = "Exchange2010_SP2"}
        elseif($radiobutton3.Checked){$Global:option = "Exchange2013_SP1"}
        elseif($radiobutton4.Checked){$Global:option = "Exchange2013_SP1"}
        $PremiseForm.Hide()
    })
    $PremiseForm.Controls.Add($buttonGo)

    #"Exit" button
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 170
    $System_Drawing_Point.Y = 50
    $buttonExit.Location = $System_Drawing_Point
    $buttonExit.Name = "Exit"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 25
    $System_Drawing_Size.Width = 50
    $buttonExit.Size = $System_Drawing_Size
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click({$PremiseForm.Close() ; $buttonExit.Dispose() })
    $PremiseForm.Controls.Add($buttonExit)

    #Show Form
    $PremiseForm.Add_Shown({$PremiseForm.Activate()})
    $PremiseForm.ShowDialog()| Out-Null
    #exit if 'Exit' button is pushed
    if($buttonExit.IsDisposed){return} 

    #creating service object
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$option
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
 
    #setting credentials
    $psCred = Get-Credential -Message "Type your credentials or Administrator credentials"
    $Global:email = $psCred.UserName
    $creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString()) 
    $service.Credentials = $creds
    $service.TraceEnabled = $False
    if($radiobutton4.Checked){
        $service.EnableScpLookup = $False 
        $service.Url = [system.URI]"https://outlook.office365.com/ews/exchange.asmx"
    }else{
        # setting Autodiscover endpoint
        $service.EnableScpLookup = $True
        $service.AutodiscoverUrl($email,{$true}) 
	}    
    Write-Host "connected to URL: " $service.url -ForegroundColor Yellow

#endregion

#region Processes

#region SelectFile Process
$SelectFileProcess= {
    $statusBar.Text = "Running..."

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.ShowDialog() | Out-Null
    if($OpenFileDialog.filename -ne ""){
        $Global:Filename = $OpenFileDialog.filename
        $txtBoxMbxAlias.Text = "...Imported from File..."
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Select file Operation finished" -ForegroundColor Yellow
        }
    $radiobutton2.Checked = $True
    $statusBar.Text = "Process Completed"
}
#endregion SelectFile Process

#region ListMeetings
$ListMeetingsProcess = {
    $statusBar.Text = "Running..." 
    if($radiobutton1.Checked){
        $mbxs = New-Object System.Object
        $mbxs | Add-Member -Type NoteProperty -Name PrimarySMTPAddress -Value $txtBoxMbxAlias.Text
    }elseif($radiobutton2.Checked){
        $mbxs = Import-Csv -Path $filename
    }
    $organizer = $TxtBoxOrganizer.Text
    $array = New-Object System.Collections.ArrayList
    $i = 0
    foreach($mbx in $mbxs){
        $i++
        $display = "Scanning mailbox " + $i + " out of " + $mbxs.count
        $txtBoxResults.Text = $display
        $txtBoxResults.Visible = $True
        $MainForm.refresh()

        $TargetSmtpAddress = $mbx.PrimarySMTPAddress
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Clear()
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        
        $Calendarfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        $startDate = $FromDatePicker.Value
        $endDate = $ToDatePicker.Value
        [int]$NumOfItems = 100
        
        $calView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, $NumOfItems)
        $calView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)

        $Appointments = $Calendarfolder.FindAppointments($calView)
        foreach ($Appointment in $Appointments){
            $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service,$Appointment.Id)
            write-host $tempItem.Subject -ForegroundColor Yellow
            if(($tempItem.Organizer.Address -like "*$organizer*") -and ($tempItem.LastModifiedName -ne $tempItem.Organizer.Name)){
                $Subject = $tempItem.subject.ToString().replace($tempItem.Organizer.Name,'')
                $output = $tempItem | Select @{N="Mailbox";E={$tempItem.LastModifiedName}},@{N="Subject";E={$Subject.trimstart()}},organizer,start,end,datetimereceived
                $array.Add($output)
            }
        }
    }
    $dgResults.datasource = $array
    $array | export-csv "$home\Desktop\ListMeetings-$organizer $((Get-Date).ToString("yyyy-MM-dd HH_mm")).csv" -NoTypeInformation
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $dgResults.AutoResizeColumns()
    $MainForm.refresh()
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - listing all meetings completed. List exported to desktop" -ForegroundColor Yellow
    $statusBar.Text = "Ready"
}
#endregion

#region DeleteMeetings
$DeleteMeetingsProcess = {
    $statusBar.Text = "Running..." 
    if($txtBoxMbxAlias.Text -ne "...Imported from File..."){
        $mbxs = New-Object System.Object
        $mbxs | Add-Member -Type NoteProperty -Name PrimarySMTPAddress -Value $txtBoxMbxAlias.Text
    }else{
        $mbxs = Import-Csv -Path $filename
    }
    $organizer = $TxtBoxOrganizer.Text

    $i = 0
    $array = New-Object System.Collections.ArrayList
    foreach($mbx in $mbxs){
        $i++
        $display = "Scanning mailbox " + $i + " out of " + $mbxs.count
        $txtBoxResults.Text = $display
        $txtBoxResults.Visible = $True
        $MainForm.refresh()

        $TargetSmtpAddress = $mbx.PrimarySMTPAddress
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetSmtpAddress)
        $service.HttpHeaders.Clear()
        $service.HttpHeaders.Add("X-AnchorMailbox", $TargetSmtpAddress)
        
        $Calendarfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        $startDate = $FromDatePicker.Value
        $endDate = $ToDatePicker.Value
        [int]$NumOfItems = 100
        
        $calView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate, $NumOfItems)
        $calView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
        
        $Appointments = $Calendarfolder.FindAppointments($calView)
        foreach ($Appointment in $Appointments){
            $tempItem = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service,$Appointment.Id)
            if(($tempItem.Organizer.Address -like "*$organizer*") -and ($tempItem.LastModifiedName -ne $tempItem.Organizer.Name)){
                $Subject = $tempItem.subject.ToString().replace($tempItem.Organizer.Name,'')
                $output = $tempItem | Select @{N="Mailbox";E={$tempItem.LastModifiedName}},@{N="Subject";E={$Subject.trimstart()}},organizer,start,end,datetimereceived
                $array.Add($output)
                $tempItem.Delete("MoveToDeletedItems","SendToNone")
            }
        }
    }
    $array | export-csv "$home\Desktop\DeletedMeetings-$organizer $((Get-Date).ToString("yyyy-MM-dd HH_mm")).csv" -NoTypeInformation
    $display = "Deletion completed. Please check your resultant file in your Desktop"
    $txtBoxResults.Text = $display
    $txtBoxResults.Visible = $True
    $MainForm.refresh()
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Deleting all meetings completed. List exported to desktop" -ForegroundColor Yellow
    $statusBar.Text = "Ready" 
}
#endregion

#endregion

#region Handlers
$OnLoadMainWindow_StateCorrection={#Correct the initial state of the form to prevent the .Net maximized form issue
	$MainForm.WindowState = $InitialFormWindowState
}

$handler_labImpersonationHelp_Click={
	[Microsoft.VisualBasic.Interaction]::MsgBox("In order to use Impersonation, we must first assign proper ManagementRole to the 'administrative' account that run the different options.
New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:<Account>

More info at: https://msdn.microsoft.com/en-us/library/bb204095(exchg.140).aspx

Press CTRL + C to copy this message to clipboard.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
}

$handler_labImportFileHowTo_Click={# Get the link to Permissions link
	[Microsoft.VisualBasic.Interaction]::MsgBox("CSV file must contain a unique header named 'PrimarySMTPAddress'.
You should list a unique Primary Email Address per line.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
}
#endregion



#----------------------------------------------
#region Generated Form Code

#Form
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Text = "Ready..."
$MainForm.Controls.Add($statusBar)
$MainForm.ClientSize = New-Object System.Drawing.Size(1000,600)
$MainForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$MainForm.Name = "form1"
$MainForm.Text = "Delete Meetings"
$MainForm.StartPosition = "CenterScreen"
$MainForm.KeyPreview = $True
$MainForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$MainForm.Close()} })
#
# radiobutton1
#
$radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$radiobutton1.Location = New-Object System.Drawing.Point(20,20)
$radiobutton1.Size = New-Object System.Drawing.Size(250,20)
$radiobutton1.TabIndex = 1
$radiobutton1.Text = "1 - Type a user/room mailbox e-mail address:"
$radioButton1.Checked = $true
$radiobutton1.UseVisualStyleBackColor = $True
$MainForm.Controls.Add($radiobutton1)
#
# txtBoxMbxAlias
#
$txtBoxMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
$txtBoxMbxAlias.Location = New-Object System.Drawing.Point(270,20)
$txtBoxMbxAlias.Size = New-Object System.Drawing.Size(150,20)
$txtBoxMbxAlias.Name = "txtBoxMbxAlias"
$MainForm.Controls.Add($txtBoxMbxAlias)
#
# radiobutton2
#
$radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$radiobutton2.Location = New-Object System.Drawing.Point(20,60)
$radiobutton2.Size = New-Object System.Drawing.Size(150,20)
$radiobutton2.TabIndex = 2
$radiobutton2.Text = "2 - import from CSV"
$radioButton2.Checked = $false
$radiobutton2.UseVisualStyleBackColor = $True
$MainForm.Controls.Add($radiobutton2)
#
# "ImportFile" button
#
$buttonImportFile.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonImportFile.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonImportFile.Location = New-Object System.Drawing.Point(270,55)
$buttonImportFile.Size = New-Object System.Drawing.Size(150,25)
$buttonImportFile.Name = "ImportFile"
$buttonImportFile.Text = ">>> Import from CSV <<<"
$buttonImportFile.UseVisualStyleBackColor = $True
$buttonImportFile.add_Click($SelectFileProcess)
$MainForm.Controls.Add($buttonImportFile)
#
# Label "File how to"
#
$labImportFileHowTo.Location = New-Object System.Drawing.Point(425,60)
$labImportFileHowTo.Size = New-Object System.Drawing.Size(50,25)
$labImportFileHowTo.Text = "?"
$labImportFileHowTo.ForeColor = "Blue"
$labImportFileHowTo.add_Click($handler_labImportFileHowTo_Click)
$MainForm.Controls.Add($labImportFileHowTo)
#
#Label FromDate
#
$labFromDate.Location = New-Object System.Drawing.Point(20,105)
$labFromDate.Size = New-Object System.Drawing.Size(120,35)
$labFromDate.Name = "FromDate"
$labFromDate.Text = "From or greater than:"
$MainForm.Controls.Add($labFromDate)
#
# FromDate Date Picker
#
$FromDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
$FromDatePicker.Location = New-Object System.Drawing.Point(270,100)
$FromDatePicker.Name = "FromDatePicker"
$FromDatePicker.Text = ""
$MainForm.Controls.Add($FromDatePicker)
#
#Label ToDate
#
$labToDate.Location = New-Object System.Drawing.Point(20,145)
$labToDate.Size = New-Object System.Drawing.Size(120,35)
$labToDate.Name = "ToDate"
$labToDate.Text = "To or less than:"
$MainForm.Controls.Add($labToDate)
#
# ToDate Date Picker
#
$ToDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
$ToDatePicker.Location = New-Object System.Drawing.Point(270,140)
$ToDatePicker.Name = "ToDatePicker"
$ToDatePicker.Text = ""
$MainForm.Controls.Add($ToDatePicker)
#
# Label Organizer
#
$labelOrganizer.Location = New-Object System.Drawing.Point(20,220)
$labelOrganizer.Size = New-Object System.Drawing.Size(160,30)
$labelOrganizer.Name = "LabelOrganizer"
$labelOrganizer.Text = "Organizer's Primary e-mail:"
$MainForm.Controls.Add($labelOrganizer)
#
# TxtBoxOrganizer
#
$TxtBoxOrganizer.DataBindings.DefaultDataSourceUpdateMode = 0
$TxtBoxOrganizer.Location = New-Object System.Drawing.Point(180,220)
$TxtBoxOrganizer.Size = New-Object System.Drawing.Size(300,20)
$TxtBoxOrganizer.Name = "TxtBoxOrganizer"
$MainForm.Controls.Add($TxtBoxOrganizer)
#
# buttonList
#
$buttonList.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonList.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonList.Location = New-Object System.Drawing.Point(700,190)
$buttonList.Size = New-Object System.Drawing.Size(100,25)
$buttonList.TabIndex = 17
$buttonList.Name = "List Meetings"
$buttonList.Text = "List Meetings"
$buttonList.UseVisualStyleBackColor = $True
$buttonList.add_Click($ListMeetingsProcess)
$MainForm.Controls.Add($buttonList)
#
# buttonDelete
#
$buttonDelete.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonDelete.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonDelete.Location = New-Object System.Drawing.Point(700,220)
$buttonDelete.Size = New-Object System.Drawing.Size(100,25)
$buttonDelete.TabIndex = 17
$buttonDelete.Name = "Delete Meetings"
$buttonDelete.Text = "Delete Meetings"
$buttonDelete.UseVisualStyleBackColor = $True
$buttonDelete.add_Click($DeleteMeetingsProcess)
$MainForm.Controls.Add($buttonDelete)
#
# Label "Help"
#
$labelHelp.Location = New-Object System.Drawing.Point(940,20)
$labelHelp.Size = New-Object System.Drawing.Size(50,25)
$labelHelp.Text = "Help Me!"
$labelHelp.ForeColor = "Blue"
$labelHelp.add_Click($handler_labImpersonationHelp_Click)
$MainForm.Controls.Add($labelHelp)
#
# "Exit" button
#
$buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonExit.Location = New-Object System.Drawing.Point(940,50)
$buttonExit.Size = New-Object System.Drawing.Size(50,25)
$buttonExit.TabIndex = 17
$buttonExit.Name = "Exit"
$buttonExit.Text = "Exit"
$buttonExit.UseVisualStyleBackColor = $True
$buttonExit.add_Click({$MainForm.Close() ; $buttonExit.Dispose() })
$MainForm.Controls.Add($buttonExit)
#
# TextBox results
#
$txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
$txtBoxResults.Location = New-Object System.Drawing.Point(5,250)
$txtBoxResults.Size = New-Object System.Drawing.Size(990,460)
$txtBoxResults.Name = "TextResults"
$txtBoxResults.BackColor = [System.Drawing.Color]::White
$txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Font = New-Object System.Drawing.Font("Consolas",8)
$txtBoxResults.Font = $Font 
$MainForm.Controls.Add($txtBoxResults)
#
#dataGrid
#
$dgResults.Anchor = 15
$dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
$dgResults.DataMember = ""
$dgResults.Location = New-Object System.Drawing.Point(5,250)
$dgResults.Size = New-Object System.Drawing.Size(990,460)
$dgResults.Name = "dgResults"
$dgResults.ReadOnly = $True
$dgResults.RowHeadersVisible = $False
$dgResults.Visible = $False
$dgResults.AllowUserToOrderColumns = $True
$dgResults.AllowUserToResizeColumns = $True
$MainForm.Controls.Add($dgResults)


#endregion Generated Form Code

# Show Form
#Save the initial state of the form
$InitialFormWindowState = $MainForm.WindowState
#Init the OnLoad event to correct the initial state of the form
$MainForm.add_Load($OnLoadMainWindow_StateCorrection)
$MainForm.Add_Shown({$MainForm.Activate()})
$MainForm.ShowDialog()| Out-Null
#exit if 'Exit' button is pushed
if($buttonExit.IsDisposed){if($EnableTranscript){stop-transcript} ; return} 
}

#Call the Function
GenerateForm
