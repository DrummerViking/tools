<#
.NOTES
	Name: Manage-Mobiles-GUI.ps1
	Authors: Agustin Gallegos
       
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.PARAMETER EnableTranscript
    Enable this parameter to write a powershell transcript in your 'Documents' folder.
.SYNOPSIS
	Allows admin to manage mobile devices with a simplified GUI
.DESCRIPTION
	Allows admin to manage mobile devices with a simplified GUI, and allow or block them in bulk
.EXAMPLE
    .\Manage-Mobiles-GUI.ps1
.EXAMPLE
    .\Manage-Mobiles-GUI.ps1 -EnableTranscript
    #>
param(
    [switch]$EnableTranscript = $false
)

$script:nl = "`r`n"
function GenerateForm {

#region Import the Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles() 
#endregion

#region Generated Form Objects
$Global:Form = New-Object System.Windows.Forms.Form
$statusBar = New-Object System.Windows.Forms.StatusBar
$radiobutton1 = New-Object System.Windows.Forms.RadioButton
$radiobutton2 = New-Object System.Windows.Forms.RadioButton
$radiobutton3 = New-Object System.Windows.Forms.RadioButton
$labelDGtitle = New-Object System.Windows.Forms.Label
$labelHelp = New-Object System.Windows.Forms.Label
$buttonGo = New-Object System.Windows.Forms.Button
$buttonExit = New-Object System.Windows.Forms.Button

$Global:dgResults = New-Object System.Windows.Forms.DataGridView
$Global:txtBoxResults = New-Object System.Windows.Forms.Label
$Global:InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

if($EnableTranscript){Start-Transcript}

#region Connecting to Exchange Online
if( (Get-PSSession).Computername -notlike "*outlook*" )
{
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) 
    {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}
#endregion

#region Processes

#region Collect Mobile Devices
$CollectMobiles = {
    $statusBar.Text = "Collecting Mobile Devices. Please wait..."
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Collecting Mobile Devices" -ForegroundColor Yellow
    #$mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"} | Select-Object SamAccountName, DisplayName, UserPrincipalName
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails "UserMailbox" -Properties SamAccountName, DisplayName, UserPrincipalName -PropertySets Minimum
    #Setting variables

    $dgResults.Visible = $False
    $txtBoxResults.Visible = $True
    $array = New-Object System.Collections.ArrayList
    $i = 0
    $mailboxCount = $mailboxes.count
    # Creating a hash table to map the Device Client Protocol
    $protocols = @{
        "REST"="REST"
        "EAS"="EAS"
        "Outlook"="HX"
    }

    #Loop through each mailbox
    foreach ($mailbox in $mailboxes) {
        $i++
        [int]$percent = $i / $mailboxCount * 100
        $output = "Scanning Users: $i out of $mailboxCount . Percent completed: $percent"
        $txtBoxResults.Text = $output
        $Form.refresh()

        $devices = Get-EXOMobileDeviceStatistics -Mailbox $mailbox.samaccountname 
        #If the current mailbox has an ActiveSync device associated, loop through each device 
        if ($devices) { 
            foreach ($device in $devices){ 
                $data = $device | Select-Object `
                @{ Name = "DisplayName" ; E={$mailbox.DisplayName}}, `
                @{ Name = "UserPrincipalName" ; E={$mailbox.UserPrincipalName}}, `
                DeviceOS, `
                DeviceID, `
                @{ Name = "ClientType" ; E={$protocols.item($device.ClientType)}}, `
                DeviceType, `
                DeviceModel, `
                DeviceAccessState, `
                DeviceAccessStateReason, `
                @{ Name = "LastSuccessSync" ; Expression = {($device.LastSuccessSync).ToString("yyyy-MM-dd HH:mm:ss")}}
                $array.Add($data) | Out-Null
            } 
        } 
    }
    $dgResults.datasource = $array
    #$dgResults.sort([System.Windows.Forms.DataGridViewColumn]$dgResults.Columns[1], "Descending")
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $dgResults.AutoResizeColumns()
    $Form.refresh()
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Collecting Mobile Devices completed" -ForegroundColor Yellow
    $statusBar.Text = "Ready..."
}
#endregion

#region ProcessAllowed
$ProcessAllowed = {
    $statusBar.Text = "Allowing selected Devices..."
    $Global:var = New-Object System.Collections.ArrayList
    foreach($row in $dgresults.selectedrows){
        $var.add($row)
        try{
            Set-CASMailbox -Identity $row.dataBoundItem.UserPrincipalName -ActiveSyncAllowedDeviceIDs @{add=$row.dataBoundItem.DeviceID}
            write-host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Successfully allowed the device" $row.dataBoundItem.DeviceID "for user" $row.dataBoundItem.UserPrincipalName -ForegroundColor Green
            $dgResults.Rows[$row.Index].DataBoundItem.DeviceAccessState = "Allowed"
            $dgResults.Rows[$row.Index].DataBoundItem.DeviceAccessStateReason = "Individual"
        }
        catch{
            write-host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Could not allowed the device" $row.dataBoundItem.DeviceID "for user" $row.dataBoundItem.UserPrincipalName -ForegroundColor Red
        }
    }
    $dgResults.AutoResizeColumns()
    $Form.refresh()
    $statusBar.Text = "Ready..."    
}
#endregion

#region ProcessBlocked
$ProcessBlocked = {
    $statusBar.Text = "Blocking selected Devices..."
    $Global:var = New-Object System.Collections.ArrayList
    foreach($row in $dgresults.selectedrows){
        $var.add($row)
        try{
            Set-CASMailbox -Identity $row.dataBoundItem.UserPrincipalName -ActiveSyncBlockedDeviceIDs @{add=$row.dataBoundItem.DeviceID}
            write-host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Successfully Blocked the device" $row.dataBoundItem.DeviceID "for user" $row.dataBoundItem.UserPrincipalName -ForegroundColor Green
            #$device = get-mobiledevice -Filter{DeviceID -eq $row.dataBoundItem.DeviceID} | select deviceaccessstate*
            $dgResults.Rows[$row.Index].DataBoundItem.DeviceAccessState = "Blocked"
            $dgResults.Rows[$row.Index].DataBoundItem.DeviceAccessStateReason = "Individual"
        }
        catch{
            write-host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Could not Blocked the device" $row.dataBoundItem.DeviceID "for user" $row.dataBoundItem.UserPrincipalName -ForegroundColor Red
        }
    }
    $dgResults.AutoResizeColumns()
    $Form.refresh()
    $statusBar.Text = "Ready..."    
}
#endregion

#endregion Processes


$handler_labelHelp_Click={
    [Microsoft.VisualBasic.Interaction]::MsgBox("Select one or more rows from the first column.
Then you can choose from the radio buttons to either allow or block the device for the user.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
}

$OnLoadMainWindow_StateCorrection={#Correct the initial state of the form to prevent the .Net maximized form issue
    $Form.WindowState = $InitialFormWindowState
    & $CollectMobiles
    $Form.refresh()
}

#----------------------------------------------
#region Generated Form Code
#
# Main Form
#
$Form.ClientSize = New-Object System.Drawing.Size(1010,550)
$Form.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$Form.Name = "form1"
$Form.Text = "Manage Mobile devices"
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $True
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Form.Close()} })
#
# Status Bar
#
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Text = "Ready..."
$Form.Controls.Add($statusBar)
#
# radiobutton1
#
$radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$radiobutton1.Location = New-Object System.Drawing.Point(20,20)
$radiobutton1.Size = New-Object System.Drawing.Size(150,15)
$radiobutton1.TabIndex = 1
$radiobutton1.Text = "1 - Allow devices"
$radioButton1.Checked = $true
$radiobutton1.UseVisualStyleBackColor = $True
$Form.Controls.Add($radiobutton1)
#
# radiobutton2
#
$radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$radiobutton2.Location = New-Object System.Drawing.Point(180,20)
$radiobutton2.Size = New-Object System.Drawing.Size(150,15)
$radiobutton2.TabIndex = 2
$radiobutton2.Text = "2 - Block devices"
$radioButton2.Checked = $false
$radiobutton2.UseVisualStyleBackColor = $True
$Form.Controls.Add($radiobutton2)
#
# radiobutton3
#
$radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$radiobutton3.Location = New-Object System.Drawing.Point(350,20)
$radiobutton3.Size = New-Object System.Drawing.Size(150,15)
$radiobutton3.TabIndex = 3
$radiobutton3.Text = "3 - option 3"
$radiobutton3.Checked = $false
$radiobutton3.UseVisualStyleBackColor = $True
#$Form.Controls.Add($radiobutton3)
#
# Label DGTitle
#
$labelDGtitle.Location = New-Object System.Drawing.Point(5,80)
$labelDGtitle.Size = New-Object System.Drawing.Size(80,20)
$labelDGtitle.Name = "Help"
$labelDGtitle.ForeColor = "Black"
$labelDGtitle.Text = "Mobile devices"
$Form.Controls.Add($labelDGtitle)
#
# Label Help
#
$labelHelp.Location = New-Object System.Drawing.Point(85,80)
$labelHelp.Size = New-Object System.Drawing.Size(10,20)
$labelHelp.Name = "Help"
$labelHelp.ForeColor = "Blue"
$labelHelp.Text = "?"
$labelHelp.add_Click($handler_labelHelp_Click)
$Form.Controls.Add($labelHelp)
#
#"Go" button
#
$buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonGo.Location = New-Object System.Drawing.Point(800,20)
$buttonGo.Size = New-Object System.Drawing.Size(50,25)
$buttonGo.TabIndex = 17
$buttonGo.Name = "Go"
$buttonGo.Text = "Go"
$buttonGo.UseVisualStyleBackColor = $True
$buttonGo.add_Click({
    if($radiobutton1.Checked){& $ProcessAllowed}
    elseif($radiobutton2.Checked){& $ProcessBlocked}
})
$Form.Controls.Add($buttonGo)
#
#"Exit" button
#
$buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonExit.Location = New-Object System.Drawing.Point(800,50)
$buttonExit.Size = New-Object System.Drawing.Size(50,25)
$buttonExit.TabIndex = 17
$buttonExit.Name = "Exit"
$buttonExit.Text = "Exit"
$buttonExit.UseVisualStyleBackColor = $True
$buttonExit.add_Click({$Form.Close() ; $buttonExit.Dispose() })
$Form.Controls.Add($buttonExit)
#
#TextBox results
#
$txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
$txtBoxResults.Location = New-Object System.Drawing.Point(5,100)
$txtBoxResults.Size = New-Object System.Drawing.Size(890,540)
$txtBoxResults.Name = "TextResults"
$txtBoxResults.BackColor = [System.Drawing.Color]::White
$txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Font = New-Object System.Drawing.Font("Consolas",8)
$txtBoxResults.Font = $Font 
$Form.Controls.Add($txtBoxResults)
#
#dataGrid
#
$dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
$dgResults.DataMember = ""
$dgResults.Location = New-Object System.Drawing.Point(5,100)
$dgResults.Size = New-Object System.Drawing.Size(1000,425)
$dgResults.Name = "dgResults"
$dgResults.ReadOnly = $True
$dgResults.SelectionMode = "FullRowSelect"
$dgResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dgResults.AllowUserToOrderColumns = $True
$dgResults.AllowUserToResizeColumns = $True
$Form.Controls.Add($dgResults)
#endregion Generated Form Code

# Show Form

#Save the initial state of the form
$InitialFormWindowState = $Form.WindowState
#Init the OnLoad event to correct the initial state of the form
$Form.add_Load($OnLoadMainWindow_StateCorrection)
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()| Out-Null

#exit if 'Exit' button is pushed
if($buttonExit.IsDisposed){if($EnableTranscript){stop-transcript} ; return} 
} #End Function
#Call the Function
GenerateForm
