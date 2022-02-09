<#
.NOTES
	Name: OnlineArchiveReport-GUI.ps1
	Authors: Agustin Gallegos

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
.SYNOPSIS
    Get reports for Mailboxes and Archives hosted in Exchange Online.

.DESCRIPTION
    Get reports for Mailboxes and Archives hosted in Exchange Online.
    Report can be viewed live in powershell interface, or send as HTML report by email. 

.EXAMPLE 
    PS C:\> OnlineArchiveReport-GUI.ps1 -EnableTranscript
    Connects to the tool and enables PowerShell Transcript

.COMPONENT
   STORE, Archive

.ROLE
   Support
#>
param(
    [switch]$EnableTranscript = $False
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

$script:nl = "`r`n"
$ProgressPreference = "SilentlyContinue"

function GenerateForm {
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
    $radiobutton3 = New-Object System.Windows.Forms.RadioButton
    $HorizontalLine = New-Object System.Windows.Forms.Label
    $labelRecipients = New-Object System.Windows.Forms.Label
    $txtBoxRecipients = New-Object System.Windows.Forms.TextBox
    $checkboxOrgAdmins = New-Object System.Windows.Forms.Checkbox
    $buttonSendEmail = New-Object System.Windows.Forms.Button

    $buttonGo = New-Object System.Windows.Forms.Button
    $buttonExit = New-Object System.Windows.Forms.Button

    $dgResults = New-Object System.Windows.Forms.DataGridView 
    $txtBoxResults = New-Object System.Windows.Forms.Label
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    #endregion Generated Form Objects

    if ($EnableTranscript) {
        Start-Transcript
    }

    #region Connect to EXO if no existing Session available
    if ( (Get-PSSession).Computername -notlike "*outlook*" ) {
        if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
            Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        }
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
    }
    #endregion

    #region Processes

    #region GetDataProcess
    $GetDataProcess = {
        $statusBar.Text = "Running..."
        $dgResults.Visible = $False
        $txtBoxResults.Visible = $True

        if ($radiobutton1.Checked) {
            $mbxs = Get-EXOmailbox -Identity $txtBoxMbxAlias.text -PropertySets Quota -ErrorAction SilentlyContinue | Select-Object UserPrincipalname, RecoverableItemsQuota
        }
        elseif ($radiobutton2.Checked) {
            $csv = Import-Csv -Path $filename
            $mbxs = $csv | ForEach-Object { get-EXOmailbox -Identity $_.UserPrincipalName -PropertySets Quota -ErrorAction SilentlyContinue | Select-Object UserPrincipalname, RecoverableItemsQuota }
        }
        elseif ($radiobutton3.Checked) {
            $mbxs = Get-EXOMailbox -ResultSize unlimited -PropertySets Quota -ErrorAction SilentlyContinue | Select-Object UserPrincipalName, RecoverableItemsQuota
        }

        if ($null -eq $mbxs) {
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox(es) doesn't have an archive associated.", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
        }
        else {
            $Global:array = New-Object System.Collections.ArrayList
            # Looping through each mailbox
            $i = 0
            if ($null -eq $mbxs.count) {
                $mbxcount = 1
            }
            else {
                $mbxcount = $mbxs.count
            }
            foreach ($mbx in $mbxs) {
                # show progess in text box
                $i++
                $j = 0
                # creating variable to store user data
                $MbxLocation = Get-MailboxLocation -User $mbx.UserPrincipalName

                Foreach ($MbxLoc in $MbxLocation) {
                    $j++
                    $output = "Checking user: $i out of: $($mbxcount)"
                    $output = $output + $nl + "Checking Mailbox object: $j out of: $($MbxLocation.count)"
                    $output = $output + $nl + "This window will refresh automatically ..."
                    $txtBoxResults.Text = $output
                    $MainForm.refresh()

                    # getting current mailbox RI quota and converting to MB
                    $global:RIQuota = $mbx | Select-Object @{Name = "RIQuota"; E = { [math]::Round(($_.RecoverableItemsQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 3) } }
                
                    #getting mailbox stats to be displayed
                    $stats = Get-EXOMailboxStatistics -Identity $MbxLoc.MailboxGuid.Guid -Properties lastlogontime | Select-Object `
                    @{N = "DisplayName"; E = { $MbxLoc.OwnerID } }, `
                    @{N = "MailboxLocationType"; E = { $MbxLoc.MailboxLocationType } }, `
                        itemcount, `
                        lastlogontime, `
                        totalitemsize, `
                    @{N = "RecoverableItemsSize"; E = { $_.totaldeleteditemsize } }, `
                    @{N = "Recoverable Items Usage Percentage"; E = { [math]::Round(($_.TotalDeletedItemSize.Value.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB) * 100 / $RIQuota.RIQuota, 3) } }
                
                    $array.Add($stats)
                }
                # sleeping process for 500 milliseconds, to prevents PS micro delays
                Start-Sleep -Milliseconds 500
            }
            $dgResults.datasource = $array
            $dgResults.AutoResizeColumns()
            $dgResults.Visible = $True
            $txtBoxResults.Visible = $False
            $MainForm.refresh()
            Clear-Variable i,j
            Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Collecting user's statistics done" -ForegroundColor Yellow
        }
        $statusBar.Text = "Ready"
    }
    #endregion

    #region SendMail
    $SendMail = {
        $statusBar.Text = "Sending E-mail..."
        [string]$html = $array | ConvertTo-Html
    
        #Replaces the HTML code with a fancier one
        $HTML = $HTML.replace('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"> <head> <title>HTML TABLE</title> </head><body>', '<html>
<style>
BODY{font-family: Arial; font-size: 8pt;}
H1{font-size: 14px;}
H2{font-size: 12px;}
H3{font-size: 12px;}
TH{border: 0px; background: #206BA4; padding: 5px; color: #EBF4FA;}
TD{border: 0px; padding: 5px; }
td.pass{background: #99CC99;}
td.passeven{background: #99CC99;}
td.warn{background: #FFCC00;}
td.fail{background: #CC0000; color: #ffffff;}
</style>
<title>Mailboxes Report</title>
<body>
<h3 style=''color:#C0C0C0;''>v1.20 (10/09/2018)</h3>
<h2 align=''left''>Mailbox Report</h2>
')
        $HTML = $HTML.Replace('</tr> <tr>', '</tr> <tr style=''background-color:#BBD9EE''>')

        $listrecipients = New-Object System.Collections.ArrayList
        #$templist = $txtBoxRecipients.text
        if ( $txtBoxRecipients.text -ne '') {
            $null = $listrecipients.Add($txtBoxRecipients.text)
        }
        # If Switch $OrgAdmins is in use, we will check current admins and include them to the recipients list
        if ($checkboxOrgAdmins.Checked -eq $True) {
            $TenantAdmins = Get-RoleGroupMember ((Get-RoleGroup tenantadmins_*).name)
            foreach ( $admin in $TenantAdmins.Name ) {
                $null = $listrecipients.Add( (Get-EXOMailbox $admin).PrimarySmtpAddress )
            }
        }
        #$listrecipients = ("$templist").Split(",")
        $Subject = "Mailbox Report $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))"
        if ($Null -eq $cred) {
            $Global:cred = Get-Credential -Message "Type your Sender's credentials"
        }
        Send-MailMessage -From $cred.UserName -To $listrecipients -Body $html -BodyasHtml -SmtpServer smtp.office365.com -UseSsl -Port 587 -Subject $Subject -Credential $cred
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - E-mail report sent" -ForegroundColor Yellow
        $statusBar.Text = "Ready"
    }
    #endregion

    #region SelectFile Process
    $SelectFileProcess = {
        $statusBar.Text = "Running..."

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.ShowDialog() | Out-Null
        if ($OpenFileDialog.filename -ne "") {
            $Global:Filename = $OpenFileDialog.filename
            $txtBoxMbxAlias.Text = "...Imported from File..."
            Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Select file Operation finished" -ForegroundColor Yellow
        }
        $radiobutton2.Checked = $True
        $statusBar.Text = "Process Completed"
    }
    #endregion SelectFile Process

    #endregion

    #region Handlers
    $OnLoadMainWindow_StateCorrection = { #Correct the initial state of the form to prevent the .Net maximized form issue
        $MainForm.WindowState = $InitialFormWindowState
    }


    $handler_labImportFileHowTo_Click = { # Get the link to Permissions link
        [Microsoft.VisualBasic.Interaction]::MsgBox("CSV file must contain a unique header named 'UserPrincipalName'.
You should list a unique UserPrincipalName per line.", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
    }
    #endregion

    #----------------------------------------------
    #region Generated Form Code

    #Form
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusStrip.name = "StatusStrip"
    $statusBar = New-Object System.Windows.Forms.ToolStripStatusLabel
    $null = $statusStrip.Items.Add($statusBar)
    $statusBar.Name = "statusBar"
    $statusBar.Text = "Ready..."
    $MainForm.Controls.Add($statusStrip)
    $MainForm.ClientSize = New-Object System.Drawing.Size(1000, 600)
    $MainForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $MainForm.Name = "form1"
    $MainForm.Text = "Online Mailbox and Archive reports"
    $MainForm.StartPosition = "CenterScreen"
    $MainForm.KeyPreview = $True
    $MainForm.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $MainForm.Close() } })
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton1.Location = New-Object System.Drawing.Point(20, 20)
    $radiobutton1.Size = New-Object System.Drawing.Size(230, 20)
    $radiobutton1.TabIndex = 1
    $radiobutton1.Text = "1 - Type the user's UserPrincipalName:"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    $MainForm.Controls.Add($radiobutton1)
    #
    # txtBoxMbxAlias
    #
    $txtBoxMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxMbxAlias.Location = New-Object System.Drawing.Point(250, 20)
    $txtBoxMbxAlias.Size = New-Object System.Drawing.Size(150, 20)
    $txtBoxMbxAlias.Name = "txtBoxMbxAlias"
    $MainForm.Controls.Add($txtBoxMbxAlias)
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20, 60)
    $radiobutton2.Size = New-Object System.Drawing.Size(150, 20)
    $radiobutton2.TabIndex = 2
    $radiobutton2.Text = "2 - import from CSV"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True
    $MainForm.Controls.Add($radiobutton2)
    #
    # "ImportFile" button
    #
    $buttonImportFile.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonImportFile.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonImportFile.Location = New-Object System.Drawing.Point(250, 55)
    $buttonImportFile.Size = New-Object System.Drawing.Size(150, 25)
    $buttonImportFile.Name = "ImportFile"
    $buttonImportFile.Text = ">>> Import from CSV <<<"
    $buttonImportFile.UseVisualStyleBackColor = $True
    $buttonImportFile.add_Click($SelectFileProcess)
    $MainForm.Controls.Add($buttonImportFile)
    #
    # Label "File how to"
    #
    $labImportFileHowTo.Location = New-Object System.Drawing.Point(405, 60)
    $labImportFileHowTo.Size = New-Object System.Drawing.Size(50, 25)
    $labImportFileHowTo.Text = "?"
    $labImportFileHowTo.ForeColor = "Blue"
    $labImportFileHowTo.add_Click($handler_labImportFileHowTo_Click)
    $MainForm.Controls.Add($labImportFileHowTo)
    #
    # radiobutton3
    #
    $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton3.Location = New-Object System.Drawing.Point(20, 100)
    $radiobutton3.Size = New-Object System.Drawing.Size(300, 15)
    $radiobutton3.TabIndex = 3
    $radiobutton3.Text = "3 - All available mailboxes in the tenant"
    $radiobutton3.Checked = $false
    $radiobutton3.UseVisualStyleBackColor = $True
    $MainForm.Controls.Add($radiobutton3)
    #
    # Horizontal Line
    #
    $HorizontalLine.Location = New-Object System.Drawing.Point(5, 130)
    $HorizontalLine.Size = New-Object System.Drawing.Size(990, 2)
    $HorizontalLine.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $HorizontalLine.Name = "Recipients"
    $HorizontalLine.Text = ""
    $MainForm.Controls.Add($HorizontalLine)
    #
    # Label Recipients
    #
    $labelRecipients.Location = New-Object System.Drawing.Point(20, 140)
    $labelRecipients.Size = New-Object System.Drawing.Size(120, 60)
    $labelRecipients.Name = "Recipients"
    $labelRecipients.Text = "Recipients to send report (separated by commas):"
    $MainForm.Controls.Add($labelRecipients)
    #
    # TxtBoxRecipients
    #
    $TxtBoxRecipients.DataBindings.DefaultDataSourceUpdateMode = 0
    $TxtBoxRecipients.Location = New-Object System.Drawing.Point(140, 140)
    $TxtBoxRecipients.Size = New-Object System.Drawing.Size(350, 20)
    $TxtBoxRecipients.Name = "TxtBoxRecipients"
    $MainForm.Controls.Add($TxtBoxRecipients)
    #
    # checkboxOrgAdmins
    #
    $checkboxOrgAdmins.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkboxOrgAdmins.Location = New-Object System.Drawing.Point(500, 140)
    $checkboxOrgAdmins.Size = New-Object System.Drawing.Size(150, 20)
    $checkboxOrgAdmins.Name = "checkboxOrgAdmins"
    $checkboxOrgAdmins.Text = "Include Tenant Admins"
    $MainForm.Controls.Add($checkboxOrgAdmins)
    #
    # buttonSendEmail
    #
    $buttonSendEmail.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonSendEmail.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonSendEmail.Location = New-Object System.Drawing.Point(700, 140)
    $buttonSendEmail.Size = New-Object System.Drawing.Size(80, 25)
    $buttonSendEmail.Name = "Send E-mail"
    $buttonSendEmail.Text = "Send E-mail"
    $buttonSendEmail.UseVisualStyleBackColor = $True
    $buttonSendEmail.add_Click($SendMail)
    $MainForm.Controls.Add($buttonSendEmail)
    #
    # "Go" button
    #
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGo.Location = New-Object System.Drawing.Point(700, 20)
    $buttonGo.Size = New-Object System.Drawing.Size(50, 25)
    $buttonGo.Name = "Go"
    $buttonGo.Text = "Run"
    $buttonGo.UseVisualStyleBackColor = $True
    $buttonGo.add_Click($GetDataProcess)
    $MainForm.Controls.Add($buttonGo)
    #
    # "Exit" button
    #
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonExit.Location = New-Object System.Drawing.Point(700, 50)
    $buttonExit.Size = New-Object System.Drawing.Size(50, 25)
    $buttonExit.Name = "Exit"
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click({ $MainForm.Close() ; $buttonExit.Dispose() })
    $MainForm.Controls.Add($buttonExit)
    #
    # TextBox results
    #
    $txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxResults.Location = New-Object System.Drawing.Point(5, 200)
    $txtBoxResults.Size = New-Object System.Drawing.Size(990, 510)
    $txtBoxResults.Name = "TextResults"
    $txtBoxResults.BackColor = [System.Drawing.Color]::White
    $txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtBoxResults.Font = $Font 
    $MainForm.Controls.Add($txtBoxResults)
    #
    #dataGrid
    #
    $dgResults.Anchor = 15
    $dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $dgResults.DataMember = ""
    $dgResults.Location = New-Object System.Drawing.Point(5, 200)
    $dgResults.Size = New-Object System.Drawing.Size(990, 510)
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
    $MainForm.Add_Shown({ $MainForm.Activate() })
    $MainForm.ShowDialog() | Out-Null
    #exit if 'Exit' button is pushed
    if ($buttonExit.IsDisposed) { if ($EnableTranscript) { stop-transcript } ; return } 
}

#Call the Function
GenerateForm

