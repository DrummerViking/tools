<#
.NOTES
	Name: Search-GUI.ps1
	Authors: Agustin Gallegos & Nelson Riera
   
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Allow admins to Search and Delete content in GUI
.DESCRIPTION
	Allow admins to Search and Delete content in GUI
#>

$script:nl = "`r`n"
$ProgressPreference = "SilentlyContinue"

function GenerateForm {

#Internal function to request inputs using UI instead of Read-Host
function Show-InputBox
{
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
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$MainWindow = New-Object System.Windows.Forms.Form
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Text = "Ready..."
$MainWindow.Controls.Add($statusBar)
$labelSearchMenu = New-Object System.Windows.Forms.Label
$labMbxAlias = New-Object System.Windows.Forms.Label
$txtBoxMbxAlias = New-Object System.Windows.Forms.TextBox
$labelDirectionName = New-Object System.Windows.Forms.Label
$comboBoxdirectionName = New-Object System.Windows.Forms.ComboBox
$labelcheckboxDumpsterOnly = New-Object System.Windows.Forms.Label
$checkboxDumpsterOnly = New-Object System.Windows.Forms.Checkbox
$labelcheckboxSoftDeleted = New-Object System.Windows.Forms.Label
$checkboxSoftDeleted = New-Object System.Windows.Forms.Checkbox
$labelcheckboxAllMbxs = New-Object System.Windows.Forms.Label
$checkboxAllMbxs = New-Object System.Windows.Forms.Checkbox
$buttonPermissions = New-Object System.Windows.Forms.Button

$labFromDate = New-Object System.Windows.Forms.Label
$FromDatePicker = New-Object System.Windows.Forms.DateTimePicker
$labToDate = New-Object System.Windows.Forms.Label
$ToDatePicker = New-Object System.Windows.Forms.DateTimePicker
$labSubject = New-Object System.Windows.Forms.Label
$txtBoxSubject = New-Object System.Windows.Forms.TextBox
$labFromFilter = New-Object System.Windows.Forms.Label
$txtBoxFromFilter = New-Object System.Windows.Forms.TextBox
$labToFilter = New-Object System.Windows.Forms.Label
$txtBoxToFilter = New-Object System.Windows.Forms.TextBox
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonImportFile = New-Object System.Windows.Forms.Button
$labImportFileHowTo = New-Object System.Windows.Forms.Label
$buttonSearchLogOnly = New-Object System.Windows.Forms.Button

$labelSearchExportMenu = New-Object System.Windows.Forms.Label
$labTargetMbxAlias = New-Object System.Windows.Forms.Label
$txtBoxTargetMbxAlias = New-Object System.Windows.Forms.TextBox
$labTargetFolder = New-Object System.Windows.Forms.Label
$txtBoxTargetFolder = New-Object System.Windows.Forms.TextBox
$buttonSearchExport = New-Object System.Windows.Forms.Button

$buttonDeleteCommand = New-Object System.Windows.Forms.Button

$labelGetRestoreRIMenu = New-Object System.Windows.Forms.Label
$labMoreInfo = New-Object System.Windows.Forms.Label
$labSourceFolderName = New-Object System.Windows.Forms.Label
$comboBoxSourceFolderName = New-Object System.Windows.Forms.ComboBox
$labItemType = New-Object System.Windows.Forms.Label
$comboBoxitemType = New-Object System.Windows.Forms.ComboBox
$buttonGetRI = New-Object System.Windows.Forms.Button
$buttonRestoreRI = New-Object System.Windows.Forms.Button

$dgResults = New-Object System.Windows.Forms.DataGridView
$txtBoxResults = New-Object System.Windows.Forms.Label
#endregion Generated Form Objects


#region connecting to powershell

# Testing if we have a live PSSession of type Exchange
$livePSSession = Get-PSSession | Where-Object{$_.ConfigurationName -eq "Microsoft.Exchange"}
if($null -ne $livePSSession){
    if($livePSSession.ComputerName -eq "outlook.office365.com"){
         $Global:premise = "office365"
    }else{
         $Global:premise = "on-premises"
        }
    }else{

    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm = New-Object System.Windows.Forms.Form
    $radiobutton1 = New-Object System.Windows.Forms.RadioButton
    $radiobutton2 = New-Object System.Windows.Forms.RadioButton
    $buttonGo = New-Object System.Windows.Forms.Button
    $buttonExit = New-Object System.Windows.Forms.Button

    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($groupbox1)
    $PremiseForm.ClientSize = New-Object System.Drawing.Size(200,100)
    $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $PremiseForm.Name = "form1"
    $PremiseForm.Text = "Choose your premises"
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton1.Location = New-Object System.Drawing.Point(20,20)
    $radiobutton1.Name = "radiobutton1"
    $radiobutton1.Size = New-Object System.Drawing.Size(100,25)
    $radiobutton1.TabStop = $True
    $radiobutton1.Text = "Office 365"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20,50)
    $radiobutton2.Name = "radiobutton2"
    $radiobutton2.Size = New-Object System.Drawing.Size(100,25)
    $radiobutton2.TabStop = $True
    $radiobutton2.Text = "On-Premises"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True

    #"Go" button
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 120
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
        if($radiobutton1.Checked){
            $Global:premise = "office365"
    }else{
         $Global:premise = "on-premises"
        }
        $PremiseForm.Close()
    })
    $PremiseForm.Controls.Add($buttonGo)

    #"Exit" button
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 120
    $System_Drawing_Point.Y = 50
    $buttonExit.Location = $System_Drawing_Point
    $buttonExit.Name = "Exit"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 25
    $System_Drawing_Size.Width = 50
    $buttonExit.Size = $System_Drawing_Size
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click({$PremiseForm.Close();$Global:premise = "exit"})
    $PremiseForm.Controls.Add($buttonExit)


    $InitialMainWindowState = $PremiseForm.WindowState
    $PremiseForm.add_Load($OnLoadMainWindow_StateCorrection)
    $PremiseForm.ShowDialog()| Out-Null

     
    if( $Global:premise -eq "exit")
    {return}
    if( $Global:premise -eq "office365"){
        if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) 
        {
            Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        }
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline        
    }else{
        # we will test common endpoints for tentative URLs based on
        # autodiscover. domain.com
        # mail .domain.com
        # webmail .domain.com
        $AutoDEmail = . Show-InputBox -Prompt "enter your E-mail Address to discover required Endpoint"
        $AutoDEmail = $AutoDEmail.Substring($AutoDEmail.IndexOf('@')+1)

        $AutoDEndpoint = $AutoDEmail.Insert(0,"autodiscover.") # definining "autodiscover.domain.com"
        if($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)){
            $AutoDEndpoint = $AutoDEmail.Insert(0,"mail.") # definining "mail.domain.com"
            if($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)){
                $AutoDEndpoint = $AutoDEmail.Insert(0,"webmail.") # definining "webmail.domain.com"
                if($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)){
                    # if all previous attempts fail, we will request to enter the Exchange Server FQDN or NETBIOS
                    $AutoDEndpoint = . Show-InputBox -Prompt "Please enter your Exchange CAS FQDN or NETBIOS name"
                }
            }
        }
        
        # Establishing session
        $Session = New-PSSession -Name Exchange -ConfigurationName Microsoft.Exchange -ConnectionUri http://$AutoDEndpoint/powershell -Authentication Kerberos -AllowRedirection
        Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue -CommandName Get-Mailbox, Search-Mailbox, Get-ManagementRoleAssignment, New-ManagementRoleAssignment, Add-RoleGroupMember, Get-RecoverableItems, Restore-RecoverableItems | Out-Null
    }
}
#endregion


#region Processes

#region Search Process
$SearchProcess= {
    $statusBar.Text = "Running..."
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    if($checkboxAllMbxs.Checked){
        $mbx = get-Mailbox -ResultSize Unlimited | Select-Object Identity
    }else{
        if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
        $csv = Import-Csv $filename
        $mbx = $csv | ForEach-Object{get-mailbox $_.primarySMTPAddress | Select-Object Identity}
        }
        if($null -ne (get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted -ErrorAction SilentlyContinue)){
            $mbx = get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted | Select-Object Identity
        }
    }

    if($null -ne $mbx){
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = "Please wait while the operation is performed."
        $output = $output + $nl + "This window will refresh automatically ..."
        $txtBoxResults.Text = $output

        $txtBoxResults.Visible = $True
        $dgResults.Visible = $False
        $MainWindow.refresh()


        # setting subject filter 

        $subject = $txtBoxSubject.Text
        if($subject -ne ""){
            $subjectFilter = "Subject: " + $subject
            }

        # setting From filter
        $fromfilter = $txtBoxFromFilter.text
        if($fromfilter -ne ""){
            $fromfilter = "From:" + $fromfilter 
        }

        
        # setting To filter
        $Tofilter = $txtBoxToFilter.text
        if($Tofilter -ne ""){
            $Tofilter = "To:" + $Tofilter 
        }

        #Dates filters
        
        $fromdate = $FromDatePicker.Value
        if($fromdate -ne ""){
            $fromdate = $direction + " >= " + $fromdate.ToString("MM-dd-yyyy")
        }
        $Todate = $ToDatePicker.Value
        if($Todate -ne ""){
            $Todate = $direction + " <= " + $Todate.ToString("MM-dd-yyyy")
        }

        if($fromdate -ne ""){
            $datefilter = $fromdate

            if($Todate -ne ""){ 
                $datefilter = $datefilter + " AND " + $Todate
                }
            }else{
            if($null -ne $Todate){ 
                $datefilter = $Todate
                }
            }

        #Combining Filters
        if($Null -ne $subjectFilter){
            $filter = $subjectFilter

            if($null -ne $datefilter){ 
                $filter = $filter + " AND " + $datefilter
                }
            }else{
            if($null -ne $datefilter){ 
                $filter = $datefilter
                }
            }
        
        if($fromfilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $fromfilter
            }else{
                $Filter = $filter + " AND " + $fromfilter
            }
        }

        if($fromfilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $Tofilter
            }else{
                $Filter = $filter + " AND " + $Tofilter
            }
        }

        $DumpsterChecked = $checkboxDumpsterOnly.Checked
        $output = $mbx | Search-Mailbox -SearchQuery $Filter -EstimateResultOnly -SearchDumpsterOnly:$DumpsterChecked -WarningAction SilentlyContinue | Select-Object Identity,Success,ResultItemsCount,ResultItemsSize,TargetMailbox,TargetFolder
        $array = New-Object System.Collections.ArrayList
        if($Null -eq $mbx.Count){
            $array.Add($output)
        }else{
            $array.addrange($output)}

	    $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $MainWindow.refresh()

        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Search Operation finished" -ForegroundColor Yellow
        $statusBar.Text = "Process Completed. Items Found: " + $output.Count
        }
         else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
            }

        #clearing Variables
        $mbx = $null
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = $null
        $DumpsterChecked = $null

}
#endregion Search Process

#region Search Export Process
$SearchExportProcess={
    $statusBar.Text = "Running..."
    if($txtBoxTargetMbxAlias.Text -eq "" -or $txtBoxTargetFolder.Text -eq ""){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Either Target Mailbox or Target Folder are empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }else{
    
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    $targetMailbox = $txtBoxTargetMbxAlias.Text
    if($null -eq (get-mailbox $targetMailbox -SoftDeletedMailbox:$IsSoftDeleted -erroraction SilentlyContinue)){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Target Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        return
        }

    if($checkboxAllMbxs.Checked){
        $mbx = get-Mailbox -ResultSize Unlimited | Select-Object Identity
    }else{
        if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
            $csv = Import-Csv $filename
            $mbx = $csv | ForEach-Object{get-mailbox $_.primarySMTPAddress | Select-Object Identity}
        }
        if($null -ne (get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted -ErrorAction SilentlyContinue)){
            $mbx = get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted | Select-Object Identity
        }
    }

    if($null -ne $mbx){

        $targetFolder = $txtBoxTargetFolder.Text
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = "Please wait while the operation is performed."
        $output = $output + $nl + "This window will refresh automatically ..."
        $txtBoxResults.Text = $output

        $txtBoxResults.Visible = $True
        $dgResults.Visible = $False
        $MainWindow.refresh()


        # setting subject filter 
        $subject = $txtBoxSubject.Text
        if($subject -ne ""){
            $subjectFilter = "Subject: " + $subject
            }


        #Dates filters
        $fromdate = $FromDatePicker.Value
        if($fromdate -ne ""){
            $fromdate = $direction + " >= " + $fromdate.ToString("MM-dd-yyyy")
        }
        $Todate = $ToDatePicker.Value
        if($Todate -ne ""){
            $Todate = $direction + " <= " + $Todate.ToString("MM-dd-yyyy")
        }

        if($fromdate -ne ""){
            $datefilter = $fromdate

            if($Todate -ne ""){ 
                $datefilter = $datefilter + " AND " + $Todate
                }
            }else{
            if($null -ne $Todate){ 
                $datefilter = $Todate
                }
            }


        #Combining Filters
        if($Null -ne $subjectFilter){
            $filter = $subjectFilter

            if($null -ne $datefilter){ 
                $filter = $filter + " AND " + $datefilter
                }
            }else{
            if($null -ne $datefilter){ 
                $filter = $datefilter
                }
            }

        # setting From filter
        $fromfilter = $txtBoxFromFilter.text
        if($fromfilter -ne ""){
            $fromfilter = "From:" + $fromfilter 
        }

        
        # setting To filter
        $Tofilter = $txtBoxToFilter.text
        if($Tofilter -ne ""){
            $Tofilter = "To:" + $Tofilter 
        }

        if($fromfilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $fromfilter
            }else{
                $Filter = $filter + " AND " + $fromfilter
            }
        }

        if($Tofilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $Tofilter
            }else{
                $Filter = $filter + " AND " + $Tofilter
            }
        }
        
        $DumpsterChecked = $checkboxDumpsterOnly.Checked
        $output = $mbx | Search-Mailbox -SearchQuery $Filter -TargetMailbox $targetMailbox -TargetFolder $targetFolder -SearchDumpsterOnly:$DumpsterChecked -WarningAction SilentlyContinue | Select-Object Identity,Success,ResultItemsCount,ResultItemsSize,TargetMailbox,TargetFolder
        $array = New-Object System.Collections.ArrayList
        if($Null -eq $mbx.Count){
            $array.Add($output)
        }else{
            $array.addrange($output)}
        
	    $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $MainWindow.refresh()

        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Search and Export Operation finished" -ForegroundColor Yellow 
        $statusBar.Text = "Process Completed. Items Found: " + $output.Count
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
            }
    }

    #clearing Variables
        $mbx = $null
        $targetMailbox = $null
        $targetFolder = $null
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = $null
        $DumpsterChecked = $null

}

#endregion Search Export Process

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
    $statusBar.Text = "Process Completed"
}
#endregion SelectFile Process

#region Search Log Only Process
$SearchLogOnlyProcess={
    $statusBar.Text = "Running..."
    if($txtBoxTargetMbxAlias.Text -eq "" -or $txtBoxTargetFolder.Text -eq ""){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Either Target Mailbox or Target Folder are empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }else{
    
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    $targetMailbox = $txtBoxTargetMbxAlias.Text
    if($null -eq (get-mailbox $targetMailbox -SoftDeletedMailbox:$IsSoftDeleted -erroraction SilentlyContinue)){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Target Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        return
        }

    if($checkboxAllMbxs.Checked){
        $mbx = get-Mailbox -ResultSize Unlimited | Select-Object Identity
    }else{
        if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
            $csv = Import-Csv $filename
            $mbx = $csv | ForEach-Object{get-mailbox $_.primarySMTPAddress | Select-Object Identity}
        }
        if($null -ne (get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted -ErrorAction SilentlyContinue)){
            $mbx = get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted | Select-Object Identity
        }
    }

    if($null -ne $mbx){

        $targetFolder = $txtBoxTargetFolder.Text
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = "Please wait while the operation is performed."
        $output = $output + $nl + "This window will refresh automatically ..."
        $txtBoxResults.Text = $output

        $txtBoxResults.Visible = $True
        $dgResults.Visible = $False
        $MainWindow.refresh()


        # setting subject filter 
        $subject = $txtBoxSubject.Text
        if($subject -ne ""){
            $subjectFilter = "Subject: " + $subject
            }


        #Dates filters
        $fromdate = $FromDatePicker.Value
        if($fromdate -ne ""){
            $fromdate = $direction + " >= " + $fromdate.ToString("MM-dd-yyyy")
        }
        $Todate = $ToDatePicker.Value
        if($Todate -ne ""){
            $Todate = $direction + " <= " + $Todate.ToString("MM-dd-yyyy")
        }

        if($fromdate -ne ""){
            $datefilter = $fromdate

            if($Todate -ne ""){ 
                $datefilter = $datefilter + " AND " + $Todate
                }
            }else{
            if($null -ne $Todate){ 
                $datefilter = $Todate
                }
            }


        #Combining Filters
        if($Null -ne $subjectFilter){
            $filter = $subjectFilter

            if($null -ne $datefilter){ 
                $filter = $filter + " AND " + $datefilter
                }
            }else{
            if($null -ne $datefilter){ 
                $filter = $datefilter
                }
            }

        # setting From filter
        $fromfilter = $txtBoxFromFilter.text
        if($fromfilter -ne ""){
            $fromfilter = "From:" + $fromfilter 
        }

        
        # setting To filter
        $Tofilter = $txtBoxToFilter.text
        if($Tofilter -ne ""){
            $Tofilter = "To:" + $Tofilter 
        }

        if($fromfilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $fromfilter
            }else{
                $Filter = $filter + " AND " + $fromfilter
            }
        }

        if($Tofilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $Tofilter
            }else{
                $Filter = $filter + " AND " + $Tofilter
            }
        }
        
        $DumpsterChecked = $checkboxDumpsterOnly.Checked
        $output = $mbx | Search-Mailbox -SearchQuery $Filter -TargetMailbox $targetMailbox -TargetFolder $targetFolder -SearchDumpsterOnly:$DumpsterChecked -LogOnly -LogLevel Full -WarningAction SilentlyContinue | Select-Object Identity,Success,ResultItemsCount,ResultItemsSize,TargetMailbox,TargetFolder
        $array = New-Object System.Collections.ArrayList
        if($Null -eq $mbx.Count){
            $array.Add($output)
        }else{
            $array.addrange($output)}
        
	    $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $MainWindow.refresh()

        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Search and Export Operation finished" -ForegroundColor Yellow 
        $statusBar.Text = "Process Completed. Items Found: " + $output.Count
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
            }
    }

    #clearing Variables
        $mbx = $null
        $targetMailbox = $null
        $targetFolder = $null
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = $null
        $DumpsterChecked = $null

}
#endregion Search Export Process

#region Delete Command Process
$DeleteCommandProcess={
    $statusBar.Text = "Running..."
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    $YesNo = $checkboxDumpsterOnly.Checked

    if($checkboxAllMbxs.Checked){
        $mbx = ""
    }else{
        $mbx = $txtBoxMbxAlias.Text
        if($txtBoxMbxAlias.Text -ne "...Imported from File..." -and
            $null -eq (get-mailbox $mbx -SoftDeletedMailbox:$IsSoftDeleted -erroraction SilentlyContinue)){
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            return
            
        }
    }
    $Filter = $null
    $subjectFilter = $null
    $datefilter = $null


    # setting subject filter 

    $subject = $txtBoxSubject.Text
    if($subject -ne ""){
        $subjectFilter = "Subject: " + $subject
        }


    #Dates filters
    $fromdate = $FromDatePicker.Value
    if($fromdate -ne ""){
        $fromdate = $direction + " >= " + $fromdate.ToString("MM-dd-yyyy")
    }
    $Todate = $ToDatePicker.Value
    if($Todate -ne ""){
        $Todate = $direction + " <= " + $Todate.ToString("MM-dd-yyyy")
    }

    if($fromdate -ne ""){
        $datefilter = $fromdate

        if($Todate -ne ""){ 
            $datefilter = $datefilter + " AND " + $Todate
            }
        }else{
        if($null -ne $Todate){ 
            $datefilter = $Todate
            }
        }

    #Combining Filters
    if($Null -ne $subjectFilter){
        $filter = $subjectFilter

        if($null -ne $datefilter){ 
            $filter = $filter + " AND " + $datefilter
            }
        }else{
        if($null -ne $datefilter){ 
            $filter = $datefilter
            }
        }

    # setting From filter
        $fromfilter = $txtBoxFromFilter.text
        if($fromfilter -ne ""){
            $fromfilter = "From:" + $fromfilter 
        }

        
        # setting To filter
        $Tofilter = $txtBoxToFilter.text
        if($Tofilter -ne ""){
            $Tofilter = "To:" + $Tofilter 
        }

        if($fromfilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $fromfilter
            }else{
                $Filter = $filter + " AND " + $fromfilter
            }
        }

        if($Tofilter -ne ""){
            if($Null -eq $Filter){
                $Filter = $Tofilter
            }else{
                $Filter = $filter + " AND " + $Tofilter
            }
        }
    
    $output = "Please verify below command. Run it in a Powershell console window." 
    $output = $output + $nl + "For security reasons, you are not allow to copy and paste. Please write it exactly as you see it below:"
    if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
        $output = $output + $nl + $nl + "`$csv = Import-Csv $filename"
        $output = $output + $nl + "`$mbx = `$csv | %{get-mailbox `$_.primarySMTPAddress | Select Identity}"
        }else{
        $output = $output + $nl + $nl + "`$mbx = Get-Mailbox $mbx -SoftDeletedMailbox:`$$IsSoftDeleted -Resultsize Unlimited | Select Identity"
        }
    $output = $output + $nl + "`$mbx | Search-Mailbox -SearchQuery `"$Filter`" -SearchDumpsterOnly:`$$YesNo -DeleteContent"
    $txtBoxResults.Text = $output
    $txtBoxResults.Visible = $true
    $dgResults.Visible = $False
    $MainWindow.refresh()

    #clearing Variables
    $mbx = $null
    $Filter = $null
    $subjectFilter = $null
    $datefilter = $null
    $output = $null
    $YesNo = $null
    $IsSoftDeleted = $null

    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Get Delete Command Operation finished" -ForegroundColor Yellow 
    $statusBar.Text = "Process Completed"
}
#endregion 

#region Get-RecoverItems Process
$GetRIProcess= {
    $statusBar.Text = "Running..."
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    if($checkboxAllMbxs.Checked){
        $mbx = get-Mailbox -ResultSize Unlimited | Select-Object Identity
    }else{
        if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
        $csv = Import-Csv $filename
        $mbx = $csv | ForEach-Object{get-mailbox $_.primarySMTPAddress | Select-Object Identity}
        }
        if($null -ne (get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted -ErrorAction SilentlyContinue)){
            $mbx = get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted | Select-Object Identity
        }
    }

    if($null -ne $mbx){
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = "Please wait while the operation is performed."
        $output = $output + $nl + "This window will refresh automatically ..."
        $txtBoxResults.Text = $output

        $txtBoxResults.Visible = $True
        $dgResults.Visible = $False
        $MainWindow.refresh()


        # setting subject filter 

        $subject = $txtBoxSubject.Text
        if($subject -ne ""){
            $SubjectContains = $subject
            }else{
            $SubjectContains = $null
            }


        #Dates filters
        
        $fromdate = $FromDatePicker.Value
        if($fromdate -ne ""){
            $FilterstartTime = $fromdate.ToString("MM/dd/yyyy")
        }

        $Todate = $ToDatePicker.Value
        if($Todate -ne ""){
            $FilterEndTime = $Todate.ToString("MM/dd/yyyy")
        }

        
        $output = $mbx | ForEach-Object{Get-RecoverableItems -Identity $_.Identity -SubjectContains $SubjectContains -FilterStartTime $FilterstartTime -FilterEndTime $FilterEndTime -FilterItemType $itemType -SourceFolder $sourceFoldername -WarningAction SilentlyContinue | Select-Object MailboxIdentity,ItemClass,Subject,lastmodifiedtime,lastParentPath,OriginalFolderExists}
        $array = New-Object System.Collections.ArrayList
        if( @($output).Count -eq 1) {
            $array.add($output)
            }else{
            $array.addrange($output) 
        }
	    $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $MainWindow.refresh()

        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Get Recoverable Items finished" -ForegroundColor Yellow
        $statusBar.Text = "Process Completed. Items Found: " + $output.Count
        }
         else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
            }
        

        #clearing Variables
        $mbx = $null
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = $null
        $DumpsterChecked = $null

}
#endregion Get-RecoverItems Process

#region Restore-RecoverItems Process
$RestoreRIProcess= {
    $statusBar.Text = "Running..."
    # creating variables
    $IsSoftDeleted = $checkboxSoftDeleted.Checked
    if($checkboxAllMbxs.Checked){
        $mbx = get-Mailbox -ResultSize Unlimited | Select-Object Identity
    }else{
        if($txtBoxMbxAlias.Text -eq "...Imported from File..."){
        $csv = Import-Csv $filename
        $mbx = $csv | ForEach-Object{get-mailbox $_.primarySMTPAddress | Select-Object Identity}
        }
        if($null -ne (get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted -ErrorAction SilentlyContinue)){
            $mbx = get-mailbox $txtBoxMbxAlias.Text -SoftDeletedMailbox:$IsSoftDeleted | Select-Object Identity
        }
    }

    if($null -ne $mbx){
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = "Please wait while the operation is performed."
        $output = $output + $nl + "This window will refresh automatically ..."
        $txtBoxResults.Text = $output

        $txtBoxResults.Visible = $True
        $dgResults.Visible = $False
        $MainWindow.refresh()


        # setting subject filter 

        $subject = $txtBoxSubject.Text
        if($subject -ne ""){
            $SubjectContains = $subject
            }else{
            $SubjectContains = $null
            }


        #Dates filters
        
        $fromdate = $FromDatePicker.Value
        if($fromdate -ne ""){
            $FilterstartTime = $fromdate.ToString("MM/dd/yyyy")
        }

        $Todate = $ToDatePicker.Value
        if($Todate -ne ""){
            $FilterEndTime = $Todate.ToString("MM/dd/yyyy")
        }

        
        $output = $mbx | ForEach-Object{Restore-RecoverableItems -Identity $_.Identity -SubjectContains $SubjectContains -FilterStartTime $FilterstartTime -FilterEndTime $FilterEndTime -FilterItemType $itemType -SourceFolder $sourceFoldername -WarningAction SilentlyContinue | Select-Object MailboxIdentity,ItemClass,Subject,RestoredtoFolderPath,WasRestored*}
        $array = New-Object System.Collections.ArrayList
        $array.addrange($output)
	    $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $MainWindow.refresh()

        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Restore Recoverable Items finished" -ForegroundColor Yellow
        $statusBar.Text = "Process Completed. Items Found: " + $output.Count
        }
         else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("Source Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            $statusBar.Text = "Process finished with warnings/errors"
            }
        

        #clearing Variables
        $mbx = $null
        $Filter = $null
        $subjectFilter = $null
        $datefilter = $null
        $output = $null
        $DumpsterChecked = $null

}
        
#endregion Restore-RecoverItems Process

#region Permissions Process
$permissionsProcess={
    $user = . Show-InputBox -Prompt "Please type the Admin's Alias you want to check permissions for"
    if($user.length -ne 0){
    if($null -eq (Get-ManagementRoleAssignment -RoleAssignee $user -Delegating $false -Role "Mailbox Search")){
        $DiscoveryMgt = 0
        }
    if($null -eq (Get-ManagementRoleAssignment -RoleAssignee $user -Delegating $false -Role "Mailbox Import Export")){
        $MbxImportExport = 0
        }
    
    if($DiscoveryMgt -eq 0 -or $MbxImportExport -eq 0){
        $result = [Microsoft.VisualBasic.Interaction]::MsgBox("You don't have all the require permissions for these operations. Do you want to add them?",[Microsoft.VisualBasic.MsgBoxStyle]::YesNo,"Permissions Missing")
        if($result -eq "yes"){
                try{
                    Add-RoleGroupMember -Identity "discovery management" -Member $user
                    write-host "User $user added to Discovery Management Role Group" -ForegroundColor Green
                }
                catch{
                    write-host "Failed to add user $user to Discovery Management Role Group" -ForegroundColor white -BackgroundColor Red
                }

                try{
                    New-ManagementRoleAssignment -Name:$user-ExportImport -Role:"Mailbox Import Export" -User:$user
                    write-host "User $user granted Mailbox Import Export permission" -ForegroundColor Green
                }
                catch{
                    write-host "Failed to grant user $user Mailbox Import Export permission" -ForegroundColor white -BackgroundColor Red
                }

                [Microsoft.VisualBasic.Interaction]::MsgBox("Operation finished. Check the Powershell window for any errors. The change might take 1 hour to apply",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
            
        }else{
            [Microsoft.VisualBasic.Interaction]::MsgBox("If you don't have the required permissions, take into account the different buttons might not work as expected.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        }

       
    }else{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Permissions are OK.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }
    }

}
#endregion


#endregion Processes


$handler_comboBoxDirectionName_SelectedIndexChanged= 
{
# Get the Event ID when item is selected
	$Global:direction = $comboBoxDirectionName.selectedItem.ToString()
}


$handler_comboBoxSourceFolderName_SelectedIndexChanged= 
{
# Get the Event ID when item is selected
	if($comboBoxSourceFolderName.selectedItem.ToString() -eq "DeletedItems/RecoverableItems"){
    $Global:SourceFoldername = $null}
    else{
    $Global:SourceFoldername = $comboBoxSourceFolderName.selectedItem.ToString()
    }
}


$handler_comboBoxitemType_SelectedIndexChanged= 
{
# Get the Event ID when item is selected
    if( $comboBoxitemType.selectedItem.ToString() -eq "All") {$Global:ItemType = $null}
	if( $comboBoxitemType.selectedItem.ToString() -eq "Mails Items"){$Global:ItemType = "IPM.Note"}
    if( $comboBoxitemType.selectedItem.ToString() -eq "Contacts Items") {$Global:ItemType = "IPM.Cibtact"}
    if( $comboBoxitemType.selectedItem.ToString() -eq "Calendar/meeting Items") {$Global:ItemType = "IPM.Appointment"}
    if( $comboBoxitemType.selectedItem.ToString() -eq "Tasks Items") {$Global:ItemType = "IPM.Task"}
}


$handler_labelSubject_Click=
{
# Get the link to Permissions link
	Start-Process -FilePath "https://msdn.microsoft.com/en-us/library/ee558911(v=office.15).aspx#Constructing free-text queries using KQL"
}


$handler_labelMoreInfo_Click=
{
# Get the link to Permissions link
	[Microsoft.VisualBasic.Interaction]::MsgBox("'Get' and 'Restore' items from Recoverable Items only works in Exchange Online.
Only 'Subject', 'Start Time' and 'End Time' filter options above will be consider for the queries.
Please select appropiate combo box options, don't leave them blank.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
}


$handler_labImportFileHowTo_Click=
{
# Get the link to Permissions link
	[Microsoft.VisualBasic.Interaction]::MsgBox("CSV file must contain a unique header named 'PrimarySMTPAddress'.
You should list a unique Primary Email Address per line.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
}

$OnLoadMainWindow_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$MainWindow.WindowState = $InitialMainWindowState
}


#----------------------------------------------
#region Generated Form Code
#main window
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 700
$System_Drawing_Size.Width = 1000
$MainWindow.ClientSize = $System_Drawing_Size
$MainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
$MainWindow.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$MainWindow.Name = "Window App"
$MainWindow.Text = "Search and Destroy App"
$MainWindow.AutoScroll = $true
$MainWindow.AutoSize = $False
$MainWindow.KeyPreview = $true
$MainWindow.Add_KeyDown({
    if($_.KeyCode -eq "Escape"){$MainWindow.Close()}
})
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$MainWindow.Icon = $Icon
$MainWindow.add_Load($handler_MainWindow_Load)


#TextBox results
$txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 340
$txtBoxResults.Location = $System_Drawing_Point
$txtBoxResults.Name = "TextResults"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 500
$System_Drawing_Size.Width = 990
$txtBoxResults.Size = $System_Drawing_Size
$txtBoxResults.BackColor = [System.Drawing.Color]::White
$txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Font = New-Object System.Drawing.Font("Consolas",8)
$txtBoxResults.Font = $Font

$MainWindow.Controls.Add($txtBoxResults)


#dataGrid

$dgResults.Anchor = 15
$dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
$dgResults.DataMember = ""
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 340
$dgResults.Location = $System_Drawing_Point
$dgResults.Name = "dgResults"
$dgResults.ReadOnly = $True
$dgResults.RowHeadersVisible = $False
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 500
$System_Drawing_Size.Width = 990
$dgResults.Size = $System_Drawing_Size
$dgResults.Visible = $False
$dgResults.AllowUserToOrderColumns = $True
$dgResults.AutoResizeColumns()

$MainWindow.Controls.Add($dgResults)



#Label "Search content on Mailbox" title

$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 5
$labelSearchMenu.Location = $System_Drawing_Point
$labelSearchMenu.Name = "Header1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$labelSearchMenu.Size = $System_Drawing_Size
$labelSearchMenu.Text = "Search content on Mailbox"
$labelSearchMenu.Font = $Font

$MainWindow.Controls.Add($labelSearchMenu)

#Label Mailbox
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 30
$labMbxAlias.Location = $System_Drawing_Point
$labMbxAlias.Name = "Mailbox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 45
$labMbxAlias.Size = $System_Drawing_Size
$labMbxAlias.Text = "Mailbox"

$MainWindow.Controls.Add($labMbxAlias)


#TextBox mailbox
$txtBoxMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 53
$System_Drawing_Point.Y = 28
$txtBoxMbxAlias.Location = $System_Drawing_Point
$txtBoxMbxAlias.Name = "txtBoxMbxAlias"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 150
$txtBoxMbxAlias.Size = $System_Drawing_Size
#By Default we will populate the user's name running the powershell
$txtBoxMbxAlias.Text = Show-InputBox -Prompt "Enter the user alias you want to check"
$MainWindow.Controls.Add($txtBoxMbxAlias)



#label Direction Name
$labelDirectionName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 30
$labelDirectionName.Location = $System_Drawing_Point
$labelDirectionName.Name = "Direction"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 60
$labelDirectionName.Size = $System_Drawing_Size
$labelDirectionName.Text = "Direction: "

$MainWindow.Controls.Add($labelDirectionName)


#ComboBox Direction Selection
$comboBoxDirectionName.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBoxDirectionName.FormattingEnabled = $True
$comboBoxDirectionName.Items.Add("")|Out-Null
$comboBoxDirectionName.Items.Add("Received")|Out-Null
$comboBoxDirectionName.Items.Add("Sent")|Out-Null
$System_Drawing_PointComboEVentID = New-Object System.Drawing.Point
$System_Drawing_PointComboEVentID.X = 270
$System_Drawing_PointComboEVentID.Y = 28
$comboBoxDirectionName.Location = $System_Drawing_PointComboEVentID
$comboBoxDirectionName.Name = "comboBoxDirectionName"
$System_Drawing_SizeComboEVentID = New-Object System.Drawing.Size
$System_Drawing_SizeComboEVentID.Height = 23
$System_Drawing_SizeComboEVentID.Width = 70
$comboBoxDirectionName.Size = $System_Drawing_SizeComboEVentID
$comboBoxDirectionName.add_SelectedIndexChanged($handler_comboBoxDirectionName_SelectedIndexChanged)
$comboBoxdirectionName.SelectedItem = "Received"
$MainWindow.Controls.Add($comboBoxDirectionName)



#label checkbox Soft Deleted
$labelcheckboxSoftDeleted.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 380
$System_Drawing_Point.Y = 30
$labelcheckboxSoftDeleted.Location = $System_Drawing_Point
$labelcheckboxSoftDeleted.Name = "SoftDeleted"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 90
$labelcheckboxSoftDeleted.Size = $System_Drawing_Size
$labelcheckboxSoftDeleted.Text = "Is Soft-Deleted:"

$MainWindow.Controls.Add($labelcheckboxSoftDeleted)

#checkbox soft Deleted
$checkboxSoftDeleted.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 470
$System_Drawing_Point.Y = 30
$checkboxSoftDeleted.Location = $System_Drawing_Point
$checkboxSoftDeleted.Name = "checkboxSoftDeleted"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 15
$checkboxSoftDeleted.Size = $System_Drawing_Size

$MainWindow.Controls.Add($checkboxSoftDeleted)



#label checkbox DumpsterOnly
$labelcheckboxDumpsterOnly.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 510
$System_Drawing_Point.Y = 30
$labelcheckboxDumpsterOnly.Location = $System_Drawing_Point
$labelcheckboxDumpsterOnly.Name = "DumpsterOnly"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 90
$labelcheckboxDumpsterOnly.Size = $System_Drawing_Size
$labelcheckboxDumpsterOnly.Text = "Dumpster Only: "

$MainWindow.Controls.Add($labelcheckboxDumpsterOnly)

#checkbox DumpsterOnly
$checkboxDumpsterOnly.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 600
$System_Drawing_Point.Y = 30
$checkboxDumpsterOnly.Location = $System_Drawing_Point
$checkboxDumpsterOnly.Name = "CheckboxDumpsterOnly"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 15
$checkboxDumpsterOnly.Size = $System_Drawing_Size

$MainWindow.Controls.Add($checkboxDumpsterOnly)



#label checkbox AllMbxs
$labelcheckboxAllMbxs.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 630
$System_Drawing_Point.Y = 30
$labelcheckboxAllMbxs.Location = $System_Drawing_Point
$labelcheckboxAllMbxs.Name = "AllMbxs"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 130
$labelcheckboxAllMbxs.Size = $System_Drawing_Size
$labelcheckboxAllMbxs.Text = "All Available Mailboxes: "

$MainWindow.Controls.Add($labelcheckboxAllMbxs)


#checkbox AllMbxs
$checkboxAllMbxs.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 761
$System_Drawing_Point.Y = 30
$checkboxAllMbxs.Location = $System_Drawing_Point
$checkboxAllMbxs.Name = "checkboxAllMbxs"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 15
$checkboxAllMbxs.Size = $System_Drawing_Size

$MainWindow.Controls.Add($checkboxAllMbxs)



#"Permissions" button
$buttonPermissions.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonPermissions.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 25
$buttonPermissions.Location = $System_Drawing_Point
$buttonPermissions.Name = "Permissions"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 200
$buttonPermissions.Size = $System_Drawing_Size
$buttonPermissions.Text = ">>> Check Admin Permissions <<<"
$buttonPermissions.UseVisualStyleBackColor = $True
$buttonPermissions.add_Click($permissionsProcess)

$MainWindow.Controls.Add($buttonPermissions)


#Label FromDate
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 57
$labFromDate.Location = $System_Drawing_Point
$labFromDate.Name = "FromDate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 80
$labFromDate.Size = $System_Drawing_Size
$labFromDate.Text = "From or greater than"

$MainWindow.Controls.Add($labFromDate)


# FromDate Date Picker
$FromDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 90
$System_Drawing_Point.Y = 57
$FromDatePicker.Location = $System_Drawing_Point
$FromDatePicker.Name = "FromDatePicker"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 70
$FromDatePicker.Text = ""
$MainWindow.Controls.Add($FromDatePicker)


#Label ToDate
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 92
$labToDate.Location = $System_Drawing_Point
$labToDate.Name = "ToDate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 40
$System_Drawing_Size.Width = 80
$labToDate.Size = $System_Drawing_Size
$labToDate.Text = "To or less than"
$MainWindow.Controls.Add($labToDate)


# ToDate Date Picker
$ToDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 90
$System_Drawing_Point.Y = 87
$ToDatePicker.Location = $System_Drawing_Point
$ToDatePicker.Name = "ToDatePicker"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 70
$ToDatePicker.Text = ""
$MainWindow.Controls.Add($ToDatePicker)


#Label Subject
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 320
$System_Drawing_Point.Y = 57
$labSubject.Location = $System_Drawing_Point
$labSubject.Name = "Subject"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 50
$labSubject.Size = $System_Drawing_Size
$labSubject.Text = "Subject: "
$labSubject.ForeColor = "Blue"
$labSubject.add_Click($handler_labelSubject_Click)

$MainWindow.Controls.Add($labSubject)


#TextBox Subject
$txtBoxSubject.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 375
$System_Drawing_Point.Y = 57
$txtBoxSubject.Location = $System_Drawing_Point
$txtBoxSubject.Name = "txtBoxToDate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 400
$txtBoxSubject.Size = $System_Drawing_Size
$txtBoxSubject.Text = ""
$MainWindow.Controls.Add($txtBoxSubject)



#Label From Filter
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 320
$System_Drawing_Point.Y = 87
$labfromFilter.Location = $System_Drawing_Point
$labfromFilter.Name = "FromFilter"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 50
$labfromFilter.Size = $System_Drawing_Size
$labfromFilter.Text = "From: "
$MainWindow.Controls.Add($labfromFilter)


#TextBox FromFilter
$txtBoxFromFilter.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 375
$System_Drawing_Point.Y = 87
$txtBoxFromFilter.Location = $System_Drawing_Point
$txtBoxFromFilter.Name = "txtBoxFromFilter"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 400
$txtBoxFromFilter.Size = $System_Drawing_Size
$txtBoxFromFilter.Text = ""
$MainWindow.Controls.Add($txtBoxFromFilter)

#Label To Filter
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 320
$System_Drawing_Point.Y = 117
$labToFilter.Location = $System_Drawing_Point
$labToFilter.Name = "FromFilter"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 50
$labToFilter.Size = $System_Drawing_Size
$labToFilter.Text = "To: "
$MainWindow.Controls.Add($labToFilter)


#TextBox ToFilter
$txtBoxToFilter.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 375
$System_Drawing_Point.Y = 117
$txtBoxToFilter.Location = $System_Drawing_Point
$txtBoxToFilter.Name = "txtBoxToFilter"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 400
$txtBoxToFilter.Size = $System_Drawing_Size
$txtBoxToFilter.Text = ""
$MainWindow.Controls.Add($txtBoxToFilter)

#"Search" button
$buttonSearch.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonSearch.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 57
$buttonSearch.Location = $System_Drawing_Point
$buttonSearch.Name = "Search"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 150
$buttonSearch.Size = $System_Drawing_Size
$buttonSearch.Text = ">>> Search <<<"
$buttonSearch.UseVisualStyleBackColor = $True
$buttonSearch.add_Click($SearchProcess)

$MainWindow.Controls.Add($buttonSearch)




#"ImportFile" button
$buttonImportFile.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonImportFile.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 87
$buttonImportFile.Location = $System_Drawing_Point
$buttonImportFile.Name = "ImportFile"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 150
$buttonImportFile.Size = $System_Drawing_Size
$buttonImportFile.Text = ">>> Import from CSV <<<"
$buttonImportFile.UseVisualStyleBackColor = $True
$buttonImportFile.add_Click($SelectFileProcess)

$MainWindow.Controls.Add($buttonImportFile)


#Label "File how to"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 940
$System_Drawing_Point.Y = 93
$labImportFileHowTo.Location = $System_Drawing_Point
$labImportFileHowTo.Name = "Subject"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 50
$labImportFileHowTo.Size = $System_Drawing_Size
$labImportFileHowTo.Text = "?"
$labImportFileHowTo.ForeColor = "Blue"
$labImportFileHowTo.add_Click($handler_labImportFileHowTo_Click)

$MainWindow.Controls.Add($labImportFileHowTo)


#"Search Log Only" button
$buttonSearchLogOnly.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonSearchLogOnly.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 117
$buttonSearchLogOnly.Location = $System_Drawing_Point
$buttonSearchLogOnly.Name = "Search Log Only"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 160
$buttonSearchLogOnly.Size = $System_Drawing_Size
$buttonSearchLogOnly.Text = ">>> Generate Log Only <<<"
$buttonSearchLogOnly.UseVisualStyleBackColor = $True
$buttonSearchLogOnly.add_Click($SearchLogOnlyProcess)

$MainWindow.Controls.Add($buttonSearchLogOnly)


#Label "Search and Export" title

$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 150
$labelSearchExportMenu.Location = $System_Drawing_Point
$labelSearchExportMenu.Name = "Header2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$labelSearchExportMenu.Size = $System_Drawing_Size
$labelSearchExportMenu.Text = "Search and Export to Target Mailbox"
$labelSearchExportMenu.Font = $Font

$MainWindow.Controls.Add($labelSearchExportMenu)


#Label Target Mailbox
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 175
$labTargetMbxAlias.Location = $System_Drawing_Point
$labTargetMbxAlias.Name = "Target Mailbox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 85
$labTargetMbxAlias.Size = $System_Drawing_Size
$labTargetMbxAlias.Text = "Target Mailbox"

$MainWindow.Controls.Add($labTargetMbxAlias)


#TextBox Target Mailbox
$txtBoxTargetMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 93
$System_Drawing_Point.Y = 172
$txtBoxTargetMbxAlias.Location = $System_Drawing_Point
$txtBoxTargetMbxAlias.Name = "txtBoxTargetMbxAlias"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 150
$txtBoxTargetMbxAlias.Size = $System_Drawing_Size
$txtBoxTargetMbxAlias.Text = ""
$MainWindow.Controls.Add($txtBoxTargetMbxAlias)


#Label Target Folder
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 293
$System_Drawing_Point.Y = 175
$labTargetFolder.Location = $System_Drawing_Point
$labTargetFolder.Name = "Target Folder"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 75
$labTargetFolder.Size = $System_Drawing_Size
$labTargetFolder.Text = "Target Folder:"

$MainWindow.Controls.Add($labTargetFolder)


#TextBox Target Folder
$txtBoxTargetFolder.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 375
$System_Drawing_Point.Y = 172
$txtBoxTargetFolder.Location = $System_Drawing_Point
$txtBoxTargetFolder.Name = "txtBoxTargetFolder"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 400
$txtBoxTargetFolder.Size = $System_Drawing_Size
$txtBoxTargetFolder.Text = ""
$MainWindow.Controls.Add($txtBoxTargetFolder)

#"Search and Export" button
$buttonSearchExport.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonSearchExport.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 172
$buttonSearchExport.Location = $System_Drawing_Point
$buttonSearchExport.Name = "SearchExport"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 180
$buttonSearchExport.Size = $System_Drawing_Size
$buttonSearchExport.Text = ">>> Search and Export <<<"
$buttonSearchExport.UseVisualStyleBackColor = $True
$buttonSearchExport.add_Click($SearchExportProcess)

$MainWindow.Controls.Add($buttonSearchExport)



#"Get Delete Command" button
$buttonDeleteCommand.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonDeleteCommand.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 785
$System_Drawing_Point.Y = 202
$buttonDeleteCommand.Location = $System_Drawing_Point
$buttonDeleteCommand.Name = "buttonDeleteCommand"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 180
$buttonDeleteCommand.Size = $System_Drawing_Size
$buttonDeleteCommand.Text = ">>> Get Delete Command <<<"
$buttonDeleteCommand.UseVisualStyleBackColor = $True
$buttonDeleteCommand.add_Click($DeleteCommandProcess)

$MainWindow.Controls.Add($buttonDeleteCommand)




#Label "Get / Restore RecoverableItems" title

$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 247
$labelGetRestoreRIMenu.Location = $System_Drawing_Point
$labelGetRestoreRIMenu.Name = "Header3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 190
$labelGetRestoreRIMenu.Size = $System_Drawing_Size
$labelGetRestoreRIMenu.Text = "Get / Restore RecoverableItems"
$labelGetRestoreRIMenu.Font = $Font

$MainWindow.Controls.Add($labelGetRestoreRIMenu)



#Label MoreInfo
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 195
$System_Drawing_Point.Y = 247
$labMoreInfo.Location = $System_Drawing_Point
$labMoreInfo.Name = "More Info"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$labMoreInfo.Size = $System_Drawing_Size
$labMoreInfo.Text = "More Info"
$labMoreInfo.ForeColor = "Blue"
$labMoreInfo.add_Click($handler_labelMoreInfo_Click)

$MainWindow.Controls.Add($labMoreInfo)


#Label Source Folder Name
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 270
$labSourceFolderName.Location = $System_Drawing_Point
$labSourceFolderName.Name = "Source Folder Name"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 150
$labSourceFolderName.Size = $System_Drawing_Size
$labSourceFolderName.Text = "Source Folder name:"

$MainWindow.Controls.Add($labSourceFolderName)


#ComboBox SourceFolder name
$comboBoxSourceFolderName.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBoxSourceFolderName.FormattingEnabled = $True
$comboBoxSourceFolderName.Items.Add("DeletedItems/RecoverableItems")|Out-Null
$comboBoxSourceFolderName.Items.Add("DeletedItems")|Out-Null
$comboBoxSourceFolderName.Items.Add("RecoverableItems")|Out-Null
$System_Drawing_PointComboEVentID = New-Object System.Drawing.Point
$System_Drawing_PointComboEVentID.X = 5
$System_Drawing_PointComboEVentID.Y = 290
$comboBoxSourceFolderName.Location = $System_Drawing_PointComboEVentID
$comboBoxSourceFolderName.Name = "comboBoxSourceFolderName"
$System_Drawing_SizeComboEVentID = New-Object System.Drawing.Size
$System_Drawing_SizeComboEVentID.Height = 23
$System_Drawing_SizeComboEVentID.Width = 180
$comboBoxSourceFolderName.Size = $System_Drawing_SizeComboEVentID
$comboBoxSourceFolderName.add_SelectedIndexChanged($handler_comboBoxSourceFolderName_SelectedIndexChanged)
$comboBoxSourceFolderName.SelectedItem = "DeletedItems/RecoverableItems"
$MainWindow.Controls.Add($comboBoxSourceFolderName)



#Label item Type
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 195
$System_Drawing_Point.Y = 270
$labItemType.Location = $System_Drawing_Point
$labItemType.Name = "Item Type"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$labItemType.Size = $System_Drawing_Size
$labItemType.Text = "Item Type:"

$MainWindow.Controls.Add($labItemType)

#ComboBox Item Type
$comboBoxitemType.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBoxitemType.FormattingEnabled = $True
$comboBoxitemType.Items.Add("All")|Out-Null
$comboBoxitemType.Items.Add("Mails Items")|Out-Null
$comboBoxitemType.Items.Add("Contacts Items")|Out-Null
$comboBoxitemType.Items.Add("Calendar/meeting Items")|Out-Null
$comboBoxitemType.Items.Add("Tasks Items")|Out-Null
$System_Drawing_PointComboEVentID = New-Object System.Drawing.Point
$System_Drawing_PointComboEVentID.X = 195
$System_Drawing_PointComboEVentID.Y = 290
$comboBoxitemType.Location = $System_Drawing_PointComboEVentID
$comboBoxitemType.Name = "comboBoxitemType"
$System_Drawing_SizeComboEVentID = New-Object System.Drawing.Size
$System_Drawing_SizeComboEVentID.Height = 23
$System_Drawing_SizeComboEVentID.Width = 130
$comboBoxitemType.Size = $System_Drawing_SizeComboEVentID
$comboBoxitemType.add_SelectedIndexChanged($handler_comboBoxitemType_SelectedIndexChanged)
$comboBoxitemType.SelectedItem = "All"
$MainWindow.Controls.Add($comboBoxitemType)

#"Get Recoverable Items" button
$buttonGetRI.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonGetRI.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 340
$System_Drawing_Point.Y = 290
$buttonGetRI.Location = $System_Drawing_Point
$buttonGetRI.Name = "buttonGetRI"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 200
$buttonGetRI.Size = $System_Drawing_Size
$buttonGetRI.Text = ">>> Get Recoverable Items <<<"
$buttonGetRI.UseVisualStyleBackColor = $True
$buttonGetRI.add_Click($GetRIProcess)

$MainWindow.Controls.Add($buttonGetRI)



#"Restore Recoverable Items" button
$buttonrestoreRI.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonrestoreRI.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 550
$System_Drawing_Point.Y = 290
$buttonrestoreRI.Location = $System_Drawing_Point
$buttonrestoreRI.Name = "buttonrestoreRI"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 200
$buttonrestoreRI.Size = $System_Drawing_Size
$buttonrestoreRI.Text = ">>> Restore Recoverable Items <<<"
$buttonrestoreRI.UseVisualStyleBackColor = $True
$buttonrestoreRI.add_Click($RestoreRIProcess)

$MainWindow.Controls.Add($buttonrestoreRI)



#endregion Generated Form Code

#Save the initial state of the form
$InitialMainWindowState = $MainWindow.WindowState
#Init the OnLoad event to correct the initial state of the form
$MainWindow.add_Load($OnLoadMainWindow_StateCorrection)
#Show the Form
$MainWindow.ShowDialog()| Out-Null
} #End Function

#Call the Function
GenerateForm
