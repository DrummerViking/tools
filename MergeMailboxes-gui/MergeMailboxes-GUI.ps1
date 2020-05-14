<#
.NOTES
	Name: MergeMailboxes-GUI.ps1
	Authors: Agustin Gallegos
                    
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
.SYNOPSIS
    Automate the process to create a New-MailboxRestoreRequest and verify the progress of it.
.DESCRIPTION
    Automate the process to create a New-MailboxRestoreRequest and verify the progress of it.
    It will allow to export SourceAccount's ProxyAddresses in case needs to be imported in the target account.
    Allows to select and combine if we involve Archive Mailboxes.
.EXAMPLE 
    .\MergeMailboxes-GUI.ps1
.COMPONENT
   STORE, Archive
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
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"

$labelSourceMailbox = New-Object System.Windows.Forms.Label
$txtBoxSourceMailbox = New-Object System.Windows.Forms.TextBox
$checkboxSourceisArchive = New-Object System.Windows.Forms.Checkbox
$labelSourceisArchive = New-Object System.Windows.Forms.Label

$labelTargetMailbox = New-Object System.Windows.Forms.Label
$txtBoxTargetMailbox = New-Object System.Windows.Forms.TextBox
$checkboxTargetisArchive = New-Object System.Windows.Forms.Checkbox
$labelTargetisArchive = New-Object System.Windows.Forms.Label

$labelBadItemLimit = New-Object System.Windows.Forms.Label
$NumericUpDown = New-Object System.Windows.Forms.NumericUpDown

$LabelListSoftDeletedMailboxes = New-Object System.Windows.Forms.Label
$labelSearchTargetMailbox = New-Object System.Windows.Forms.Label
$TxtBoxSearchTargetMailbox = New-Object System.Windows.Forms.TextBox

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonGo = New-Object System.Windows.Forms.Button
$buttonExit = New-Object System.Windows.Forms.Button

$dgResults = New-Object System.Windows.Forms.DataGridView 
$dgResults2 = New-Object System.Windows.Forms.DataGridView 
$txtBoxResults = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects
 
if($EnableTranscript){Start-Transcript}

#region Connect to EXO if no existing Session available
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

#region Collect Soft-Deleted mailboxes
$statusBar.Text = "Collecting Soft-Deleted mailboxes..."
$mbxs = Get-EXOMailbox -ResultSize Unlimited -softDeletedMailbox -PropertySets StatisticsSeed, archive | Select-Object UserPrincipalName,ExchangeGuid,ArchiveGuid
$array = New-Object System.Collections.ArrayList
foreach($mbx in $mbxs){
    $array.Add($mbx) | Out-Null
}
$dgResults.datasource = $array
$dgResults.Visible = $True
$txtBoxResults.Visible = $False
$dgResults.AutoResizeColumns()
$MainForm.refresh()
Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Collecting Soft-Deleted mailboxes completed" -ForegroundColor Yellow
$statusBar.Text = "Ready"
#endregion

#region Search Target mailbox
$ProcessSearchTargetMailbox = {
    $statusBar.Text = "Searching Target mailbox..."
    $Global:targetMailbox = Get-EXOMailbox -Anr $TxtBoxSearchTargetMailbox.text -PropertySets StatisticsSeed, Archive | Select-Object UserPrincipalName,ExchangeGuid,ArchiveGuid
    $array = New-Object System.Collections.ArrayList
    foreach($mbx in $targetMailbox){
        $array.Add($mbx) | Out-Null
    }
    $dgResults2.datasource = $array
    $dgResults2.Visible = $True
    $txtBoxResults.Visible = $False
    $dgResults2.AutoResizeColumns()
    $MainForm.refresh()
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Searching target mailbox completed" -ForegroundColor Yellow
    $statusBar.Text = "Ready"
}
#endregion

#region Merge
$ProcessMerge = {
    $statusBar.Text = "Merging mailboxes..."

    # getting Addresses from soft-deleted account
    $sourceMailbox = Get-EXOMailbox $txtBoxSourceMailbox.Text -SoftDeletedMailbox -PropertySets StatisticsSeed, archive, minimum | Select-Object UserPrincipalName,ExchangeGuid,emailAddresses,ArchiveGuid

    # removing SIP and SPO addresses from the Array
    foreach($address in $sourceMailbox.EmailAddresses){
        if( ($address).StartsWith("SPO",[System.StringComparison]::CurrentCulture)){
            $address1 = $address
            }
        if( ($address).StartsWith("SIP",[System.StringComparison]::CurrentCulture)){
            $address2 = $address
            }
    }
    ($sourceMailbox.EmailAddresses).remove($address1)
    ($sourceMailbox.EmailAddresses).remove($address2)

    # Saving Guids for later operations
    if($checkboxSourceisArchive.Checked){ $SourceGuid = $sourceMailbox.ArchiveGuid ; $sourceSwitch = $True
    }else{$SourceGuid = $sourceMailbox.ExchangeGuid ; $sourceSwitch = $False}

    if($checkboxTargetisArchive.Checked){ $targetGuid = $targetMailbox.ArchiveGuid ; $targetSwitch = $True
    }else{$targetGuid = $targetMailbox.ExchangeGuid ; $targetSwitch = $False}

    # starting new merge
    $name = "Restore Request " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    New-MailboxRestoreRequest -Name $name -SourceMailbox $SourceGuid.Guid -SourceIsArchive:$sourceSwitch -TargetMailbox $targetGuid.Guid -TargetIsArchive:$targetSwitch -AllowLegacyDNMismatch -BadItemLimit $NumericUpDown.Value -AcceptLargeDataLoss -WarningAction SilentlyContinue | Out-Null
    $output = "Creating new Mailbox Restore Request" + $nl
    $output += "New-MailboxRestoreRequest -Name `"$name`" -SourceMailbox $SourceGuid -SourceIsArchive:`$$sourceSwitch -TargetMailbox $targetGuid -TargetIsArchive:`$$targetSwitch -AllowLegacyDNMismatch -BadItemLimit $($NumericUpDown.Value) -AcceptLargeDataLoss" + $nl
    $output += "Status will appear shortly"
    $txtBoxResults.Text = $output
    $MainForm.refresh()
    
    $req = Get-MailboxRestoreRequest | Where-Object {$_.Identity -like "*$name*"}
    $txtBoxResults.Visible = $True
    while((Get-MailboxRestoreRequest | Where-Object {$_.Identity -like "*$name*"}).Status -ne "Completed"){
        $RequestStatus = Get-MailboxRestoreRequestStatistics $req.RequestGuid
        $output = "Status " + $RequestStatus.StatusDetail.Value + ", PercentComplete " + $RequestStatus.PercentComplete + $nl
        $output += "ItemsTransferred: " + $RequestStatus.ItemsTransferred + ", Percent Completed: " + $RequestStatus.PercentComplete
        Start-Sleep -Milliseconds 500
    
        $txtBoxResults.Text = $output
        $MainForm.refresh()
    }
    $output = "Mailbox Merged finished"
    $txtBoxResults.Text = $output
    $MainForm.refresh()
    
    # opting the user to save or not proxyAddresses from disconnected mailbox to TXT file in the desktop
    $opt = Show-InputBox -Prompt "Please type 'Y' if you want to save the E-mail Addresses from the disconnected mailbox to a text file in your desktop or press Cancel"
    switch($opt){
        Y {
            $filename = "$env:userprofile\desktop\ProxyAddresses.txt"
            Write-Host "Creating file $filename" -ForegroundColor Green
            $stream = [System.IO.StreamWriter] "$filename"
            $stream.Write("list of Proxy Addresses exported from Cloud's Disconnected Mailbox
-------------------------------------------------------------------
Add these addresses to the target user account. Be aware that if the user is Synchronized from Active Directory, you need to do it on-premises."
)
            foreach($address in $sourceMailbox.EmailAddresses){
    $stream.Writeline("$address")
            }
            
            $stream.Close()
          }
    }
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Merging mailboxes completed" -ForegroundColor Yellow
    $statusBar.Text = "Ready"
}
#endregion

#endregion

#region Handlers
$OnLoadMainWindow_StateCorrection={#Correct the initial state of the form to prevent the .Net maximized form issue
	$MainForm.WindowState = $InitialFormWindowState
}
#endregion

#----------------------------------------------
#region Generated Form Code

#Form
$statusBar.Text = "Ready..."
$MainForm.Controls.Add($statusBar)
$MainForm.ClientSize = New-Object System.Drawing.Size(1000,800)
$MainForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
$MainForm.Name = "form1"
$MainForm.Text = "Online Mailbox Merge"
$MainForm.StartPosition = "CenterScreen"
$MainForm.KeyPreview = $True
$MainForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$MainForm.Close()} })

#
# Label SourceMailbox
#
$labelSourceMailbox.Location = New-Object System.Drawing.Point(20,22)
$labelSourceMailbox.Size = New-Object System.Drawing.Size(120,30)
$labelSourceMailbox.Name = "SourceMailbox"
$labelSourceMailbox.Text = "Source Mailbox Guid:"
$MainForm.Controls.Add($labelSourceMailbox)
#
# TxtBox SourceMailbox
#
$TxtBoxSourceMailbox.DataBindings.DefaultDataSourceUpdateMode = 0
$TxtBoxSourceMailbox.Location = New-Object System.Drawing.Point(145,20)
$TxtBoxSourceMailbox.Size = New-Object System.Drawing.Size(350,20)
$TxtBoxSourceMailbox.Name = "TxtBoxSourceMailbox"
$MainForm.Controls.Add($TxtBoxSourceMailbox)
#
# Checkbox SourceisArchive
#
$checkboxSourceisArchive.DataBindings.DefaultDataSourceUpdateMode = 0
$checkboxSourceisArchive.Location = New-Object System.Drawing.Point(500,20)
$checkboxSourceisArchive.Size = New-Object System.Drawing.Size(15,20)
$checkboxSourceisArchive.Name = "checkboxSourceisArchive"
$MainForm.Controls.Add($checkboxSourceisArchive)
#
# Label SourceisArchive
#
$labelSourceisArchive.Location = New-Object System.Drawing.Point(520,22)
$labelSourceisArchive.Size = New-Object System.Drawing.Size(120,30)
$labelSourceisArchive.Name = "SourceisArchive"
$labelSourceisArchive.Text = "Source is Archive"
$MainForm.Controls.Add($labelSourceisArchive)

#
# Label TargetMailbox
#
$labelTargetMailbox.Location = New-Object System.Drawing.Point(20,52)
$labelTargetMailbox.Size = New-Object System.Drawing.Size(120,30)
$labelTargetMailbox.Name = "TargetMailbox"
$labelTargetMailbox.Text = "Target Mailbox Guid:"
$MainForm.Controls.Add($labelTargetMailbox)
#
# TxtBox TargetMailbox
#
$TxtBoxTargetMailbox.DataBindings.DefaultDataSourceUpdateMode = 0
$TxtBoxTargetMailbox.Location = New-Object System.Drawing.Point(145,50)
$TxtBoxTargetMailbox.Size = New-Object System.Drawing.Size(350,20)
$TxtBoxTargetMailbox.Name = "TxtBoxTargetMailbox"
$MainForm.Controls.Add($TxtBoxTargetMailbox)
#
# Checkbox TargetisArchive
#
$checkboxTargetisArchive.DataBindings.DefaultDataSourceUpdateMode = 0
$checkboxTargetisArchive.Location = New-Object System.Drawing.Point(500,50)
$checkboxTargetisArchive.Size = New-Object System.Drawing.Size(15,20)
$checkboxTargetisArchive.Name = "checkboxSourceisArchive"
$MainForm.Controls.Add($checkboxTargetisArchive)
#
# Label TargetisArchive
#
$labelTargetisArchive.Location = New-Object System.Drawing.Point(520,52)
$labelTargetisArchive.Size = New-Object System.Drawing.Size(120,30)
$labelTargetisArchive.Name = "TargetisArchive"
$labelTargetisArchive.Text = "Target is Archive"
$MainForm.Controls.Add($labelTargetisArchive)

#
# Label BadItemLimit
#
$labelBadItemLimit.Location = New-Object System.Drawing.Point(20,82)
$labelBadItemLimit.Size = New-Object System.Drawing.Size(100,30)
$labelBadItemLimit.Name = "BadItemLimit"
$labelBadItemLimit.Text = "Bad Item Limit:"
$MainForm.Controls.Add($labelBadItemLimit)
#
# TxtBox NumericUpDown
#
$NumericUpDown.DataBindings.DefaultDataSourceUpdateMode = 0
$NumericUpDown.Location = New-Object System.Drawing.Point(145,80)
$NumericUpDown.Size = New-Object System.Drawing.Size(40,30)
$NumericUpDown.Name = "TxtBoxBadItemLimit"
$NumericUpDown.Minimum = 1
$NumericUpDown.Maximum = 999
$MainForm.Controls.Add($NumericUpDown)

#
# Label ListSoftDeletedMailboxes
#
$LabelListSoftDeletedMailboxes.Location = New-Object System.Drawing.Point(5,130)
$LabelListSoftDeletedMailboxes.Size = New-Object System.Drawing.Size(160,30)
$LabelListSoftDeletedMailboxes.Name = "LabelListSoftDeletedMailboxes"
$LabelListSoftDeletedMailboxes.Text = "List of Soft-Deleted Mailboxes:"
$MainForm.Controls.Add($LabelListSoftDeletedMailboxes)
#
# Label SearchTargetMailbox
#
$LabelSearchTargetMailbox.Location = New-Object System.Drawing.Point(500,120)
$LabelSearchTargetMailbox.Size = New-Object System.Drawing.Size(120,30)
$LabelSearchTargetMailbox.Name = "LabelSearchTargetMailbox"
$LabelSearchTargetMailbox.Text = "SearchTarget Mailbox:"
$MainForm.Controls.Add($LabelSearchTargetMailbox)
#
# TxtBox SearchTargetMailbox
#
$TxtBoxSearchTargetMailbox.DataBindings.DefaultDataSourceUpdateMode = 0
$TxtBoxSearchTargetMailbox.Location = New-Object System.Drawing.Point(620,118)
$TxtBoxSearchTargetMailbox.Size = New-Object System.Drawing.Size(230,20)
$TxtBoxSearchTargetMailbox.Name = "TxtBoxSearchTargetMailbox"
$MainForm.Controls.Add($TxtBoxSearchTargetMailbox)
#
# "Go" button
#
$buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonGo.Location = New-Object System.Drawing.Point(880,20)
$buttonGo.Size = New-Object System.Drawing.Size(55,25)
$buttonGo.TabIndex = 17
$buttonGo.Name = "Go"
$buttonGo.Text = "Merge"
$buttonGo.UseVisualStyleBackColor = $True
$buttonGo.add_Click($ProcessMerge)
$MainForm.Controls.Add($buttonGo)
#
# "Exit" button
#
$buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonExit.Location = New-Object System.Drawing.Point(880,50)
$buttonExit.Size = New-Object System.Drawing.Size(50,25)
$buttonExit.TabIndex = 17
$buttonExit.Name = "Exit"
$buttonExit.Text = "Exit"
$buttonExit.UseVisualStyleBackColor = $True
$buttonExit.add_Click({$MainForm.Close() ; $buttonExit.Dispose() })
$MainForm.Controls.Add($buttonExit)
#
# "Search" button
#
$buttonSearch.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonSearch.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$buttonSearch.Location = New-Object System.Drawing.Point(880,118)
$buttonSearch.Size = New-Object System.Drawing.Size(60,25)
$buttonSearch.TabIndex = 17
$buttonSearch.Name = "Search"
$buttonSearch.Text = "Search"
$buttonSearch.UseVisualStyleBackColor = $True
$buttonSearch.add_Click($ProcessSearchTargetMailbox)
$MainForm.Controls.Add($buttonSearch)
#
# TextBox results
#
$txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
$txtBoxResults.Location = New-Object System.Drawing.Point(5,675)
$txtBoxResults.Size = New-Object System.Drawing.Size(990,125)
$txtBoxResults.Name = "TextResults"
$txtBoxResults.BackColor = [System.Drawing.Color]::White
$txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Font = New-Object System.Drawing.Font("Consolas",8)
$txtBoxResults.Font = $Font 
$txtBoxResults.Visible = $True
$MainForm.Controls.Add($txtBoxResults)
#
#dataGrid
#
$dgResults.Anchor = 15
$dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
$dgResults.DataMember = ""
$dgResults.Location = New-Object System.Drawing.Point(5,160)
$dgResults.Size = New-Object System.Drawing.Size(490,510)
$dgResults.Name = "dgResults"
$dgResults.ReadOnly = $True
$dgResults.RowHeadersVisible = $False
$dgResults.Visible = $True
$dgResults.AllowUserToOrderColumns = $True
$dgResults.AllowUserToResizeColumns = $True
$MainForm.Controls.Add($dgResults)
#
#dataGrid2
#
$dgResults2.Anchor = 15
$dgResults2.DataBindings.DefaultDataSourceUpdateMode = 0
$dgResults2.DataMember = ""
$dgResults2.Location = New-Object System.Drawing.Point(495,160)
$dgResults2.Size = New-Object System.Drawing.Size(490,510)
$dgResults2.Name = "dgResults"
$dgResults2.ReadOnly = $True
$dgResults2.RowHeadersVisible = $False
$dgResults2.Visible = $True
$dgResults2.AllowUserToOrderColumns = $True
$dgResults2.AllowUserToResizeColumns = $True
$MainForm.Controls.Add($dgResults2)
#endregion Generated Form Code

# Show Form
#Save the initial state of the form
$InitialFormWindowState = $MainForm.WindowState
#Init the OnLoad event to correct the initial state of the form
$MainForm.add_Load($OnLoadMainWindow_StateCorrection)
$MainForm.Add_Shown({$MainForm.Activate()})
$MainForm.ShowDialog()| Out-Null
$MainForm.Refresh()
#exit if 'Exit' button is pushed
if($buttonExit.IsDisposed){if($EnableTranscript){stop-transcript} ; return} 
}

#Call the Function
GenerateForm