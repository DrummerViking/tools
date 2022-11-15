<#
.NOTES
	Name: Manage-FolderPermissionsGUI.ps1
	Author: Agustin Gallegos
    
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Allow admins to manage a user's folder permissions in a GUI, withouth the need of an Outlook client, or powershell scripting knowledge.
.DESCRIPTION
	This file loads a GUI (Powershell Forms) to allow an admin to manage their user's mailbox folder permissions. It allows to add, remove and get permissions.
    It has a simple logic to try to connect to on-premises environments automatically.
    It has been tested in Exchange 2013 and Office 365.
#>
function GenerateForm {

    #Internal function to request inputs using UI instead of Read-Host
    function Show-InputBox {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $Prompt,
        
            [Parameter(Mandatory = $false)]
            [string]
            $DefaultValue = '',
        
            [Parameter(Mandatory = $false)]
            [string]
            $Title = 'Windows PowerShell'
        )
    
    
        Add-Type -AssemblyName Microsoft.VisualBasic
        [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $DefaultValue)
    }

    #region Import the Assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
    [System.Windows.Forms.Application]::EnableVisualStyles() 
    #endregion

    #region Generated Form Objects
    $MainWindow = New-Object System.Windows.Forms.Form
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusStrip.name = "StatusStrip"
    $statusBar = New-Object System.Windows.Forms.ToolStripStatusLabel
    $null = $statusStrip.Items.Add($statusBar)
    $statusBar.Name = "statusBar"
    $statusBar.Text = "Ready..."
    $MainWindow.Controls.Add($statusStrip)
    $labAssignMbxPermHeader = New-Object System.Windows.Forms.Label
    $labMbxAlias = New-Object System.Windows.Forms.Label
    $txtBoxMbxAlias = New-Object System.Windows.Forms.TextBox
    $handler_MailboxOwner_changed = New-Event System.Windows.Controls.TextChangedEventHandler($txtBoxMbxAlias, $null)
    $labelFolderName = New-Object System.Windows.Forms.Label
    $comboBoxFolderName = New-Object System.Windows.Forms.ComboBox
    $buttonGo = New-Object System.Windows.Forms.Button
    $buttonGo3 = New-Object System.Windows.Forms.Button
    $labTo = New-Object System.Windows.Forms.Label
    $txtBoxTo = New-Object System.Windows.Forms.TextBox
    $labelPermission = New-Object System.Windows.Forms.Label
    $comboBoxMbxPermissions = New-Object System.Windows.Forms.ComboBox

    $labGetMbxPermHeader = New-Object System.Windows.Forms.Label
    $labelFolderName_GetPerm = New-Object System.Windows.Forms.Label
    $comboBoxFolderName_GetPerm = New-Object System.Windows.Forms.ComboBox
    $buttonGo2 = New-Object System.Windows.Forms.Button

    $dgResults = New-Object System.Windows.Forms.DataGridView
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    #endregion Generated Form Objects

    #region connecting to powershell
    # Testing if we have a live PSSession of type Exchange
    $livePSSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
    if ($null -ne $livePSSession) {
        if ($livePSSession.ComputerName -eq "outlook.office365.com") {
            $premise = "Office365"
        }
        else {
            $premise = "on-premises"
        }
    }
    else {
        # Selecting if connection is to Office 365 or an Exchange on-premises
        $premise = . Show-InputBox -Prompt "Type 'Office365' if your mailbox is online or 'on-premises' if you have a local Exchange.
Type 'exit' to quit"
        while ($premise -ne "Office365" -and $premise -ne "on-premises" -and $premise -ne "exit") {
            $premise = . Show-InputBox -Prompt "Please try again. Type 'Office365' if your mailbox is online or 'on-premises' if you have a local Exchange
Type 'exit' to quit"
        }
        switch ($premise) {
            default { return }
            office365 {
                if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable) ) {
                    Install-Module ExchangeOnlineManagement -scope CurrentUser -Force -ErrorAction Stop
                }
                Import-Module ExchangeOnlineManagement
                get-connectionInformation
                try {
                    Connect-ExchangeOnline -ErrorAction Stop
                }
                catch {
                    if ($_.Exception.Message -match "8856f961-340a-11d0-a96b-00c04fd705a2") {
                        Connect-ExchangeOnline -ShowBanner:$False
                    }
                    else {
                        write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Something failed to connect. Error message: $_" -ForegroundColor Red
                    }
                }
            }
        
            on-premises {
                $handler_comboBoxAuthOpt_SelectedIndexChanged = {
                    # Get the Event ID when item is selected
                    $Global:ComboOption = $Global:comboBoxAuthOpt.selectedItem.ToString()
                }
            
                # creating GUI objects for this request
                $labText = New-Object System.Windows.Forms.Label
                $Global:comboBoxAuthOpt = New-Object System.Windows.Forms.ComboBox
                $Global:txtBoxURL = New-Object System.Windows.Forms.TextBox
                $buttonGo = New-Object System.Windows.Forms.Button
                $buttonExit = New-Object System.Windows.Forms.Button

                #main window
                $MainWindow.ClientSize = New-Object System.Drawing.Size(450, 230)
                $MainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
                $MainWindow.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
                $MainWindow.Name = "Window App"
                $MainWindow.Text = "Select manual URL"
                $MainWindow.KeyPreview = $true
                $MainWindow.Add_KeyDown( {
                    if ($_.KeyCode -eq "Escape") { $MainWindow.Close() }
                })
                $MainWindow.add_Load($handler_MainWindow_Load)
            
                $labText.Location = New-Object System.Drawing.Point(5, 5)
                $labText.Size = New-Object System.Drawing.Size(300, 163)
                $labText.Name = "labText"
                $labText.Text = "Please insert your URL.
If you are in your internal network it might be: 
http://yourCASServername/powershell

If you are connecting from the internet, it might be: 
https://mail.contoso.com/powershell
            
Choose your Authentication methods.
Tipically from your internal network it can be: Basic, Kerberos, Negotiate, NegotiateWithImplicitCredential
From the internet is tipically Basic

URL:"
                $MainWindow.Controls.Add($labText)

                $txtBoxURL.Location = New-Object System.Drawing.Point(5, 175)
                $txtBoxURL.Size = New-Object System.Drawing.Size(280, 20)
                $txtBoxURL.Name = "txtBoxURL"
                $txtBoxURL.Text = ""
                $MainWindow.Controls.Add($txtBoxURL)

                $comboBoxAuthOpt.DataBindings.DefaultDataSourceUpdateMode = 0
                $comboBoxAuthOpt.FormattingEnabled = $True
                $comboBoxAuthOpt.Location = New-Object System.Drawing.Point(320, 175)
                $comboBoxAuthOpt.Name = "comboBoxAuthOpt"
                $comboBoxAuthOpt.Items.Add("") | Out-Null
                $comboBoxAuthOpt.Items.Add("Basic") | Out-Null
                $comboBoxAuthOpt.Items.Add("Kerberos") | Out-Null
                $comboBoxAuthOpt.Items.Add("Negotiate") | Out-Null
                $comboBoxAuthOpt.Items.Add("NegotiateWithImplicitCredential") | Out-Null
                $comboBoxAuthOpt.add_SelectedIndexChanged($handler_comboBoxAuthOpt_SelectedIndexChanged)
                $MainWindow.Controls.Add($comboBoxAuthOpt)


                #"Go" button
                $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
                $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
                $buttonGo.Location = New-Object System.Drawing.Point(380, 20)
                $buttonGo.Size = New-Object System.Drawing.Size(50, 25)
                $buttonGo.Name = "buttonGo"
                $buttonGo.Text = "Go"
                $buttonGo.UseVisualStyleBackColor = $True
                $buttonGo.add_Click( {
                        $MainWindow.Controls.RemoveByKey("labText")
                        $MainWindow.Controls.RemoveByKey("comboBoxAuthOpt")
                        $MainWindow.Controls.RemoveByKey("txtBoxURL")
                        $MainWindow.Controls.RemoveByKey("buttonGo")
                        $MainWindow.Controls.RemoveByKey("buttonExit")
                        $MainWindow.Hide()
                    })
                $MainWindow.Controls.Add($buttonGo)

                #"Exit" button
                $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
                $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
                $buttonExit.Location = New-Object System.Drawing.Point(380, 50)
                $buttonExit.Size = New-Object System.Drawing.Size(50, 25)
                $buttonExit.Name = "buttonExit"
                $buttonExit.Text = "Exit"
                $buttonExit.UseVisualStyleBackColor = $True
                $buttonExit.add_Click( { $MainWindow.Close() ; $buttonExit.Dispose() })
                $MainWindow.Controls.Add($buttonExit)

                #Show MainWindow
                $MainWindow.Add_Shown( { $MainWindow.Activate() })
                $MainWindow.ShowDialog() | Out-Null
                #exit if 'Exit' button is pushed
                if ($buttonExit.IsDisposed) { return } 

                # Establishing session
                $Session = New-PSSession -Name Exchange -ConfigurationName Microsoft.Exchange -ConnectionUri $txtBoxURL.Text -Authentication $ComboOption -AllowRedirection -Credential (Get-Credential)
                Import-PSSession $Session -AllowClobber -CommandName Get-Mailbox, Get-MailboxFolderStatistics, Add-MailboxFolderPermission, Get-MailboxFolderPermission, Remove-MailboxFolderPermission -WarningAction SilentlyContinue | Out-Null
            }   
        }
    }
    #endregion


    #region Processes

    #region Process to Add Folder Permission
    $processData = 
    {
        $statusBar.Text = "Running..."
        $mbxTo = $txtBoxTo.Text
        $mbxAlias = $txtBoxMbxAlias.Text
        $mbxAlias += ":"
        $user = get-mailbox $mbxTo -ErrorAction SilentlyContinue
        if ($Null -eq $user) {
            $txtBoxTo.Text = Show-InputBox -Prompt "We can't find that user by the way you typed it. please try again and hit the 'Run' button"
        }
        else {
            if ($ChoiceFolderName1 -eq "APPLY TO ALL FOLDERS") {
                $folders = get-MailboxFolderStatistics $txtBoxMbxAlias.Text | Select-Object FolderPath
                foreach ($folder in $folders) {
                    $folderpath = $folder.FolderPath -replace "/", "\"
                    Add-MailboxFolderPermission -Identity "$mbxAlias$folderpath" -User $user.Name -AccessRights $ChoiceMbxPermission -ErrorAction SilentlyContinue
                    write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Permission '$ChoiceMbxPermission' for user '$user' added in folder '$mbxAlias$folderpath'" -ForegroundColor white -BackgroundColor DarkGreen
                }
            }
            else {
                $folderpath = $ChoiceFolderName1 -replace "/", "\"
                Add-MailboxFolderPermission -Identity "$mbxAlias$folderpath" -User $user.Name -AccessRights $ChoiceMbxPermission -ErrorAction SilentlyContinue
                write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Permission '$ChoiceMbxPermission' for user '$user' added in folder '$mbxAlias$folderpath'" -ForegroundColor white -BackgroundColor DarkGreen
            }
        
        }
        $MainWindow.refresh()
    
        #clearing variables
        $mbxTo = $null
        $mbxAlias = $null
        $folderpath = $null
        $user = $null
        $statusBar.Text = "Process Completed"
    
    }
    #endregion

    #region Process to just list FolderPermissions in a desired folder
    $processData2 =
    {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList
        $mbxAlias = $txtBoxMbxAlias.Text
        $folderpath = $ChoiceFolderName2 -replace "/", "\"
        $mbxAlias += ":"
        $output = Get-MailboxFolderPermission "$mbxAlias$folderpath" | Select-Object FolderName, User, AccessRights
        if ($null -eq $output.User.count) {
            $array.add($output) | Out-Null
            $dgResults.datasource = $array
            $MainWindow.refresh()
        }
        else {
            $array.addrange($output) | Out-Null
            $dgResults.datasource = $array
            $MainWindow.refresh()
        }
        write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Getting permissions from folder $mbxAlias$folderpath" -ForegroundColor white -BackgroundColor DarkGray

        $array = $null
        $mbxAlias = $null
        $folderpath = $null
        $output = $null
        $statusBar.Text = "Process Completed"

    }
    #endregion

    #region Process to Remove Folder Permission
    $processData3 = 
    {
        $statusBar.Text = "Running..."
        $mbxTo = $txtBoxTo.Text
        $mbxAlias = $txtBoxMbxAlias.Text
        $mbxAlias += ":"
        $user = get-mailbox $mbxTo -ErrorAction SilentlyContinue
        if ($Null -eq $user) {
            $txtBoxTo.Text = Show-InputBox -Prompt "We can't find that user by the way you typed it. please try again and hit the 'Remove' button"
        }
        else {
            if ($ChoiceFolderName1 -eq "APPLY TO ALL FOLDERS") {
                $folders = get-MailboxFolderStatistics $txtBoxMbxAlias.Text | Select-Object FolderPath
                foreach ($folder in $folders) {
                    $folderpath = $folder.FolderPath -replace "/", "\"
                    $perm = Get-MailboxFolderPermission -Identity "$mbxAlias$folderpath" -User $user.Name -ErrorAction SilentlyContinue
                    if ($Null -ne $perm) {
                        $perm | Remove-MailboxFolderPermission -Confirm:$false -User $user.Name
                        write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Permission '$ChoiceMbxPermission' for user '$user' removed in folder '$mbxAlias$folderpath'" -ForegroundColor white -BackgroundColor Red
                    }
                }
            }
            else {
                $folderpath = $ChoiceFolderName1 -replace "/", "\"
                $perm = Get-MailboxFolderPermission -Identity "$mbxAlias$folderpath" -User $user.Name -ErrorAction SilentlyContinue
                if ($Null -ne $perm) {
                    $perm | Remove-MailboxFolderPermission -Confirm:$false -User $user.Name
                    write-host "$((Get-Date).ToString("MM/dd/yyyy HH:mm:ss")) - Permission '$ChoiceMbxPermission' for user '$user' removed in folder '$mbxAlias$folderpath'" -ForegroundColor white -BackgroundColor Red
                }
            }
        }
        $MainWindow.refresh()

        #clearing variables
        $mbxTo = $null
        $mbxAlias = $null
        $folderpath = $null
        $user = $null
        $statusBar.Text = "Process Completed"
    }
    #endregion
    #endregion


    $handler_comboBoxFolderName_GetPerm_SelectedIndexChanged = 
    {
        # Get the Folder selected to get permissions from
        $Global:ChoiceFolderName2 = $comboBoxFolderName_GetPerm.selectedItem.ToString()
    }

    $handler_comboBoxFolderName_SelectedIndexChanged = 
    {
        # Get the Folder selected to assign permissions on to
        $Global:ChoiceFolderName1 = $comboBoxFolderName.selectedItem.ToString()
    }

    $handler_comboBoxMbxPermissions_SelectedIndexChanged = 
    {
        # Get the Mailbox Permission when item is selected
        $Global:ChoiceMbxPermission = $comboBoxMbxPermissions.selectedItem.ToString()
    }

    $handler_labelPermission_Click =
    {
        # Get the link to Permissions link
        Start-Process -FilePath "https://technet.microsoft.com/en-us/library/ff522363(v=exchg.160).aspx"
    }

    $handler_MailboxOwner_changed =
    {
        # clearing comboBoxes
        $comboBoxFolderName.Items.Clear()
        $comboBoxFolderName_GetPerm.Items.Clear()

        # Update ComboBoxes
        $folderstats = get-MailboxFolderStatistics $txtBoxMbxAlias.Text

        $comboBoxFolderName.Items.Add("APPLY TO ALL FOLDERS") | Out-Null
        foreach ($folder in $folderstats) {
            $comboBoxFolderName.Items.Add($folder.FolderPath) | Out-Null
            $comboBoxFolderName_GetPerm.Items.Add($folder.FolderPath) | Out-Null
        }
    }

    $OnLoadMainWindow_StateCorrection =
    { #Correct the initial state of the form to prevent the .Net maximized form issue
        $MainWindow.WindowState = $InitialMainWindowState
    }

    #----------------------------------------------
    #region Generated Form Code
    #main window
    $MainWindow.ClientSize = New-Object System.Drawing.Size(1000, 400)
    $MainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
    $MainWindow.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $MainWindow.Name = "Window App"
    $MainWindow.Text = "Managing Mailbox Folders Permissions"
    $MainWindow.KeyPreview = $true
    $MainWindow.Add_KeyDown( {
            if ($_.KeyCode -eq "Escape") { $MainWindow.Close() }
        })
    Start-Transcript -Path "$home\desktop\Permissions log $((Get-Date).ToString("MM-dd-yyyy HH_mm_ss")).txt" -Force
    $MainWindow.add_Load($handler_MainWindow_Load)
    #
    #dataGrid
    #
    $dgResults.Anchor = 15
    $dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $dgResults.DataMember = ""
    $dgResults.Location = New-Object System.Drawing.Point(3, 200)
    $dgResults.Size = New-Object System.Drawing.Size(990, 500)
    $dgResults.Name = "dgResults"
    $dgResults.ReadOnly = $True
    $dgResults.RowHeadersVisible = $false    
    $dgResults.AllowUserToOrderColumns = $True
    $dgResults.AllowUserToResizeColumns = $True
    $MainWindow.Controls.Add($dgResults)
    #
    #Label "Assigning Mailbox Permissions" title
    #
    $Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labAssignMbxPermHeader.Location = New-Object System.Drawing.Point(3, 5)
    $labAssignMbxPermHeader.Size = New-Object System.Drawing.Size(200, 20)
    $labAssignMbxPermHeader.Name = "Header1"
    $labAssignMbxPermHeader.Text = "Assigning Mailbox Permissions"
    $labAssignMbxPermHeader.Font = $Font
    $MainWindow.Controls.Add($labAssignMbxPermHeader)
    #
    #Label Mailbox Owner
    #
    $labMbxAlias.Location = New-Object System.Drawing.Point(3, 30)
    $labMbxAlias.Size = New-Object System.Drawing.Size(87, 20)
    $labMbxAlias.Name = "Mailbox"
    $labMbxAlias.Text = "Manage Mailbox"
    $MainWindow.Controls.Add($labMbxAlias)
    #
    #TextBox mailbox Owner
    #
    $txtBoxMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxMbxAlias.Location = New-Object System.Drawing.Point(90, 28)
    $txtBoxMbxAlias.Size = New-Object System.Drawing.Size(150, 20)
    $txtBoxMbxAlias.Name = "txtBoxMbxAlias"
    #By Default we will populate the user's name running the powershell
    $txtBoxMbxAlias.Text = Show-InputBox -Prompt "Enter the user alias you want to manage"
    #$txtBoxMbxAlias.TextChanged = $handler_MailboxOwner_changed
    $MainWindow.Controls.Add($txtBoxMbxAlias)
    #
    #label Folder Name
    #
    $labelFolderName.DataBindings.DefaultDataSourceUpdateMode = 0
    $labelFolderName.Location = New-Object System.Drawing.Point(257, 30)
    $labelFolderName.Size = New-Object System.Drawing.Size(50, 23)
    $labelFolderName.Name = "Folder"
    $labelFolderName.Text = "Folder: "
    $MainWindow.Controls.Add($labelFolderName)
    #
    #ComboBox FolderName Selection
    #
    $comboBoxFolderName.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxFolderName.FormattingEnabled = $True
    $comboBoxFolderName.Location = New-Object System.Drawing.Point(317, 28)
    $comboBoxFolderName.Size = New-Object System.Drawing.Size(400, 23)
    $folderstats = get-MailboxFolderStatistics $txtBoxMbxAlias.Text
    $comboBoxFolderName.Items.Add("APPLY TO ALL FOLDERS") | Out-Null
    foreach ($folder in $folderstats) {
        $comboBoxFolderName.Items.Add($folder.FolderPath) | Out-Null
    }
    $comboBoxFolderName.Name = "comboBoxFolderName"
    $comboBoxFolderName.add_SelectedIndexChanged($handler_comboBoxFolderName_SelectedIndexChanged)
    $MainWindow.Controls.Add($comboBoxFolderName)
    #
    #"run" button
    #
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGo.Location = New-Object System.Drawing.Point(762, 28)
    $buttonGo.Size = New-Object System.Drawing.Size(150, 25)
    $buttonGo.Name = "button1"
    $buttonGo.Text = ">>> Add <<<"
    $buttonGo.UseVisualStyleBackColor = $True
    $buttonGo.add_Click($processData)
    $MainWindow.Controls.Add($buttonGo)
    #
    #"Remove" button
    #
    $buttonGo3.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo3.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGo3.Location = New-Object System.Drawing.Point(762, 57)
    $buttonGo3.Size = New-Object System.Drawing.Size(150, 25)
    $buttonGo3.Name = "button3"
    $buttonGo3.Text = ">>> Remove <<<"
    $buttonGo3.UseVisualStyleBackColor = $True
    $buttonGo3.add_Click($processData3)
    $MainWindow.Controls.Add($buttonGo3)
    #
    #Label User To grant access
    #
    $labTo.Location = New-Object System.Drawing.Point(3, 57)
    $labTo.Size = New-Object System.Drawing.Size(115, 20)
    $labTo.Name = "labTo"
    $labTo.Text = "Grant permissions to:"
    $MainWindow.Controls.Add($labTo)
    #
    #TextBox User To grant access
    #
    $txtBoxTo.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxTo.Location = New-Object System.Drawing.Point(120, 55)
    $txtBoxTo.Size = New-Object System.Drawing.Size(250, 20)
    $txtBoxTo.Name = "txtBoxTo"
    $txtBoxTo.Text = ""
    $MainWindow.Controls.Add($txtBoxTo)
    #
    #Label Permission listing
    #
    $labelPermission.DataBindings.DefaultDataSourceUpdateMode = 0
    $labelPermission.Location = New-Object System.Drawing.Point(385, 57)
    $labelPermission.Size = New-Object System.Drawing.Size(100, 23)
    $labelPermission.Name = "Folder"
    $labelPermission.Text = "Permission Role: "
    $labelPermission.ForeColor = "Blue"
    $labelPermission.add_Click($handler_labelPermission_Click)
    $MainWindow.Controls.Add($labelPermission)
    #
    #ComboBox Permission Selection
    #
    $comboBoxMbxPermissions.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxMbxPermissions.FormattingEnabled = $True
    $comboBoxMbxPermissions.Location = New-Object System.Drawing.Point(485, 55)
    $comboBoxMbxPermissions.Size = New-Object System.Drawing.Size(100, 23)
    $comboBoxMbxPermissions.Items.Add("") | Out-Null
    $comboBoxMbxPermissions.Items.Add("Author") | Out-Null
    $comboBoxMbxPermissions.Items.Add("Contributor") | Out-Null
    $comboBoxMbxPermissions.Items.Add("Editor") | Out-Null
    $comboBoxMbxPermissions.Items.Add("None") | Out-Null
    $comboBoxMbxPermissions.Items.Add("NonEditingAuthor") | Out-Null
    $comboBoxMbxPermissions.Items.Add("Owner") | Out-Null
    $comboBoxMbxPermissions.Items.Add("PublishingEditor") | Out-Null
    $comboBoxMbxPermissions.Items.Add("PublishingAuthor") | Out-Null
    $comboBoxMbxPermissions.Items.Add("Reviewer") | Out-Null
    $comboBoxMbxPermissions.Name = "comboBoxMbxPermissions"
    $comboBoxMbxPermissions.add_SelectedIndexChanged($handler_comboBoxMbxPermissions_SelectedIndexChanged)
    $MainWindow.Controls.Add($comboBoxMbxPermissions)
    #
    #Label "Getting Mailbox Permissions" title
    #
    $Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labGetMbxPermHeader.Location = New-Object System.Drawing.Point(3, 140)
    $labGetMbxPermHeader.Size = New-Object System.Drawing.Size(200, 20)
    $labGetMbxPermHeader.Name = "Header2"
    $labGetMbxPermHeader.Text = "Getting Mailbox Permissions"
    $labGetMbxPermHeader.Font = $Font
    $MainWindow.Controls.Add($labGetMbxPermHeader)
    #
    #label Folder Name for Get Permissions
    #
    $labelFolderName_GetPerm.DataBindings.DefaultDataSourceUpdateMode = 0
    $labelFolderName_GetPerm.Location = New-Object System.Drawing.Point(3, 165)
    $labelFolderName_GetPerm.Size = New-Object System.Drawing.Size(100, 30)
    $labelFolderName_GetPerm.Name = "Folder"
    $labelFolderName_GetPerm.Text = "Folder to display permissions: "
    $MainWindow.Controls.Add($labelFolderName_GetPerm)
    #
    #ComboBox FolderName Selection for Get Permissions
    #
    $comboBoxFolderName_GetPerm.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBoxFolderName_GetPerm.FormattingEnabled = $True
    $comboBoxFolderName_GetPerm.Location = New-Object System.Drawing.Point(118, 163)
    $comboBoxFolderName_GetPerm.Size = New-Object System.Drawing.Size(400, 23)
    $folderstats = get-MailboxFolderStatistics $txtBoxMbxAlias.Text
    foreach ($folder in $folderstats) {
        $comboBoxFolderName_GetPerm.Items.Add($folder.FolderPath) | Out-Null
    }
    $comboBoxFolderName_GetPerm.Name = "comboBoxFolderName_GetPerm"
    $comboBoxFolderName_GetPerm.add_SelectedIndexChanged($handler_comboBoxFolderName_GetPerm_SelectedIndexChanged)
    $MainWindow.Controls.Add($comboBoxFolderName_GetPerm)
    #
    #"Get" button
    #
    $buttonGo2.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo2.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGo2.Location = New-Object System.Drawing.Point(600, 163)
    $buttonGo2.Size = New-Object System.Drawing.Size(150, 25)
    $buttonGo2.Name = "button2"
    $buttonGo2.Text = ">>> Get <<<"
    $buttonGo2.UseVisualStyleBackColor = $True
    $buttonGo2.add_Click($processData2)
    $MainWindow.Controls.Add($buttonGo2)
    #endregion Generated Form Code

    #Save the initial state of the form
    $InitialMainWindowState = $MainWindow.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $MainWindow.add_Load($OnLoadMainWindow_StateCorrection)
    #Show the Form
    $MainWindow.ShowDialog() | Out-Null
} #End Function

#Call the Function
GenerateForm

Stop-Transcript -ErrorAction SilentlyContinue