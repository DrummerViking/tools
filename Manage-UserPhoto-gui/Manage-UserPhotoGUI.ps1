<#
.NOTES
	Name: Manage-UserPhotoGUI.ps1
	Authors: Agustin Gallegos

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Allow admins to upload user Photos to Exchange Online using a GUI
.DESCRIPTION
	Allow admins to upload user Photos to Exchange Online using a GUI.
    We grant the option to create a RBAC Role Group, with the minimum permissions to list mailboxes and manage UserPhotos. This is intended for a help desk assignment.
#>

$script:nl = "`r`n"

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
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
#  
#################################################################################
"@
Write-Host $disclaimer -foregroundColor Yellow
Write-Host " " 


function GenerateForm { 
#region Import the Assemblies
Add-Type -AssemblyName Microsoft.VisualBasic
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion
 
#region Generated Form Objects
$MainWindow = New-Object System.Windows.Forms.Form
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Text = "Ready..."
$MainWindow.Controls.Add($statusBar)

$labelMenu = New-Object System.Windows.Forms.Label
$pictureBox = new-object Windows.Forms.PictureBox
$buttonSelectFile = New-Object System.Windows.Forms.Button

$labelMailbox = New-Object System.Windows.Forms.Label
$txtBoxMailbox = New-Object System.Windows.Forms.TextBox
$buttonUpload = New-Object System.Windows.Forms.Button
$buttonRemove = New-Object System.Windows.Forms.Button

$labelRoleGroupMenu = New-Object System.Windows.Forms.Label
$labUserToAdd = New-Object System.Windows.Forms.Label
$txtBoxUserToAdd = New-Object System.Windows.Forms.TextBox
$buttonUserToAdd = New-Object System.Windows.Forms.Button
$buttonUserToRemove = New-Object System.Windows.Forms.Button
$buttonCreateRoleGroup = New-Object System.Windows.Forms.Button
$buttonGetRoleGroup = New-Object System.Windows.Forms.Button

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects
 
#region connecting to powershell

# Testing if we have a live PSSession of type Exchange
$livePSSession = Get-PSSession | Where-Object{$_.ConfigurationName -eq "Microsoft.Exchange"}
if($null -ne $livePSSession){
    if($livePSSession.ComputerName -eq "outlook.office365.com"){
        $premise = "Office365"
    }else{
        $premise = "on-premises"
        }
    }else{
    # Choosing if connection is to Office 365 or an Exchange on-premises
    $premise = . Show-InputBox -Prompt "Type 'Office365' if your Admin account is online or 'on-premises' if you have a local Exchange.
Type 'exit' to quit"
    while($premise -ne "Office365" -and $premise -ne "on-premises" -and $premise -ne "exit"){
	    $premise = . Show-InputBox -Prompt "Please try again. Type 'Office365' if your Admin account is online or 'on-premises' if you have a local Exchange
Type 'exit' to quit"
    }
    if($premise -eq "exit")
    {return}
    if($premise -eq "office365"){
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
        $null = Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue
    }
}     
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
        $Image = [System.Drawing.Image]::Fromfile($filename)
        $pictureBox.Image = $Image.GetThumbnailImage(240,240,$null,0)
        $MainWindow.Controls.Add($pictureBox)
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Select file Operation finished" -ForegroundColor Yellow
        }
    $statusBar.Text = "Process Completed"
}
#endregion SelectFile Process

#region Upload Process
$UploadProcess={
    $statusBar.Text = "Running. Please wait..."
    if($txtBoxMailbox.Text -eq "" -or $Null -eq $filename){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox or Picture box is empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }elseif($null -ne (Get-Mailbox $txtBoxMailbox.Text -ErrorAction SilentlyContinue)){    
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Uploading Photo to $($txtBoxMailbox.text). Please wait" -ForegroundColor Yellow 
        Set-UserPhoto $txtBoxMailbox.Text -PictureData ([System.IO.File]::ReadAllBytes($filename)) -Confirm:$false -Preview
        Set-UserPhoto $txtBoxMailbox.Text -Save -Confirm:$false
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Uploaded Photo to $($txtBoxMailbox.text) Operation finished" -ForegroundColor Green 
        
    }else{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }
    $filename = $null
    $statusBar.Text = "Process Completed"
    
}
#endregion Upload Process

#region remove Process
$RemoveProcess={
    $statusBar.Text = "Running. Please wait..."
    if($txtBoxMailbox.Text -eq ""){
        [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox box is empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }elseif($null -ne (Get-Mailbox $txtBoxMailbox.Text -ErrorAction SilentlyContinue)){    
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Removing Photo to $($txtBoxMailbox.text). Please wait" -ForegroundColor Yellow 
        remove-UserPhoto $txtBoxMailbox.Text -Confirm:$false
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Removed Photo to $($txtBoxMailbox.text) Operation finished" -ForegroundColor Green 
        
    }else{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }
    Clear-Variable $filename -ErrorAction SilentlyContinue
    $statusBar.Text = "Process Completed"
    
}
#endregion Remove Process

#region CreateRoleGroup Process
$CreateRoleGroupProcess={
    $statusBar.Text = "Running. Please wait..."
    if($Null -ne (Get-ManagementRole "UserPhoto Roles" -ErrorAction SilentlyContinue) -or $Null -ne (get-RoleGroup "UserPhoto Admins" -ErrorAction SilentlyContinue)){
            [Microsoft.VisualBasic.Interaction]::MsgBox("Groups already exists.",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        }else{
            new-ManagementRole -Parent "Mail Recipients" -Name "UserPhoto - Mail Recipients" -EnabledCmdlets Set-UserPhoto,Remove-UserPhoto
            new-ManagementRole -Parent "View-Only Recipients" -Name "UserPhoto - View-Only Recipients" -EnabledCmdlets Get-Mailbox
            New-RoleGroup "UserPhoto Admins" -Roles "UserPhoto - Mail Recipients","UserPhoto - View-Only Recipients"
            }
    $statusBar.Text = "Process Completed"
    
}
#endregion

#region AddUsertoRoleProcess Process
$AddUsertoRoleProcess={
    $statusBar.Text = "Running. Please wait..."
    if($txtBoxUserToAdd.Text -eq ""){
        [Microsoft.VisualBasic.Interaction]::MsgBox("User box is empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }else{    
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Adding user $($txtBoxUserToAdd.text) to management Role 'UserPhoto Admins'." -ForegroundColor Yellow 
        Add-RoleGroupMember -Identity "UserPhoto Admins" -Member $txtBoxUserToAdd.Text
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Successfully added user $($txtBoxUserToAdd.text) to management Role 'UserPhoto Admins'." -ForegroundColor Green 
        
    $statusBar.Text = "Process Completed"
    
    }
}
#endregion

#region RemoveUsertoRoleProcess Process
$RemoveUserfromRoleProcess={
    $statusBar.Text = "Running. Please wait..."
    if($txtBoxUserToAdd.Text -eq ""){
        [Microsoft.VisualBasic.Interaction]::MsgBox("User box is empty. Please check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
    }else{    
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Removing user $($txtBoxUserToAdd.text) from management Role 'UserPhoto Admins'." -ForegroundColor Yellow 
        Remove-RoleGroupMember -Identity "UserPhoto Admins" -Member $txtBoxUserToAdd.Text -Confirm:$false
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Successfully removed user $($txtBoxUserToAdd.text) from management Role 'UserPhoto Admins'." -ForegroundColor Green 
        
    $statusBar.Text = "Process Completed"
    
    }
}
#endregion



#region GetRoleGroupProcess Process
$GetRoleGroupProcess={
    $statusBar.Text = "Running. Please wait..."
    Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Getting Role 'UserPhoto Admins' members." -ForegroundColor Yellow 
    try
    {
        Get-RoleGroupMember "UserPhoto Admins" -ErrorAction Stop | Out-Host
        Write-Host "$((Get-Date).ToString("MM-dd-yyyy HH:mm:ss")) - Getting Role 'UserPhoto Admins' members finished." -ForegroundColor Yellow
    }
    catch
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("'UserPhoto Admins' Role group is not created yet or something failed",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
    }        
    $statusBar.Text = "Process Completed"
}
#endregion

#endregion Processes


$OnLoadMainWindow_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$MainWindow.WindowState = $InitialMainWindowState
}


#----------------------------------------------
#region Generated Form Code
#main window
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 350
$System_Drawing_Size.Width = 1000
$MainWindow.ClientSize = $System_Drawing_Size
$MainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
$MainWindow.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$MainWindow.Name = "Window App"
$MainWindow.Text = "UserPhoto App"
$MainWindow.AutoScroll = $true
$MainWindow.AutoSize = $False
$MainWindow.KeyPreview = $true
$MainWindow.Add_KeyDown({
    if($_.KeyCode -eq "Escape"){$MainWindow.Close()}
})
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$MainWindow.Icon = $Icon
$MainWindow.add_Load($handler_MainWindow_Load)
 

#Label Menu

$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 5
$labelMenu.Location = $System_Drawing_Point
$labelMenu.Name = "Menu"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$labelMenu.Size = $System_Drawing_Size
$labelMenu.Text = "Manage User Photo"
$labelMenu.Font = $Font
 
$MainWindow.Controls.Add($labelMenu)
 

#Picture box
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 25
$pictureBox.Location = $System_Drawing_Point
$pictureBox.Width =  240
$pictureBox.Height =  240
$MainWindow.Controls.Add($pictureBox)



#"Select File" button
$buttonSelectFile.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonSelectFile.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 3
$System_Drawing_Point.Y = 275
$buttonSelectFile.Location = $System_Drawing_Point
$buttonSelectFile.Name = "SelectFile"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 150
$buttonSelectFile.Size = $System_Drawing_Size
$buttonSelectFile.Text = "Select File"
$buttonSelectFile.UseVisualStyleBackColor = $True
$buttonSelectFile.add_Click($SelectFileProcess)

$MainWindow.Controls.Add($buttonSelectFile)


#Label Mailbox
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 275
$System_Drawing_Point.Y = 30
$labelMailbox.Location = $System_Drawing_Point
$labelMailbox.Name = "Mailbox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 50
$labelMailbox.Size = $System_Drawing_Size
$labelMailbox.Text = "Mailbox: "
 
$MainWindow.Controls.Add($labelMailbox)
 
 
#TextBox Mailbox
$txtBoxMailbox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 330
$System_Drawing_Point.Y = 30
$txtBoxMailbox.Location = $System_Drawing_Point
$txtBoxMailbox.Name = "txtBoxMailbox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 400
$txtBoxMailbox.Size = $System_Drawing_Size
$txtBoxMailbox.Text = ""
$MainWindow.Controls.Add($txtBoxMailbox)


#"Upload" button
$buttonUpload.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonUpload.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 30
$buttonUpload.Location = $System_Drawing_Point
$buttonUpload.Name = "Upload"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 150
$buttonUpload.Size = $System_Drawing_Size
$buttonUpload.Text = ">>> Upload <<<"
$buttonUpload.UseVisualStyleBackColor = $True
$buttonUpload.add_Click($UploadProcess)

$MainWindow.Controls.Add($buttonUpload)


#"Remove" button
$buttonRemove.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonRemove.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 60
$buttonRemove.Location = $System_Drawing_Point
$buttonRemove.Name = "Remove"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 150
$buttonRemove.Size = $System_Drawing_Size
$buttonRemove.Text = ">>> Remove <<<"
$buttonRemove.UseVisualStyleBackColor = $True
$buttonRemove.add_Click($RemoveProcess)

$MainWindow.Controls.Add($buttonRemove)



#Label labelRoleGroupMenu
$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 275
$System_Drawing_Point.Y = 150
$labelRoleGroupMenu.Location = $System_Drawing_Point
$labelRoleGroupMenu.Name = "Role Group Menu"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$labelRoleGroupMenu.Size = $System_Drawing_Size
$labelRoleGroupMenu.Text = "Manage Role Group"
$labelRoleGroupMenu.Font = $Font
 
$MainWindow.Controls.Add($labelRoleGroupMenu)

#Label User to Add
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 275
$System_Drawing_Point.Y = 175
$labUserToAdd.Location = $System_Drawing_Point
$labUserToAdd.Name = "User to Add"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 75
$labUserToAdd.Size = $System_Drawing_Size
$labUserToAdd.Text = "User to add:"
 
$MainWindow.Controls.Add($labUserToAdd)
 
 
#TextBox User to Add
$txtBoxUserToAdd.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 350
$System_Drawing_Point.Y = 172
$txtBoxUserToAdd.Location = $System_Drawing_Point
$txtBoxUserToAdd.Name = "txtBoxUserToAdd"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 380
$txtBoxUserToAdd.Size = $System_Drawing_Size
$txtBoxUserToAdd.Text = ""
$MainWindow.Controls.Add($txtBoxUserToAdd)


#"Add to Role" button
$buttonUserToAdd.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonUserToAdd.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 172
$buttonUserToAdd.Location = $System_Drawing_Point
$buttonUserToAdd.Name = "buttonUserToAdd"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 180
$buttonUserToAdd.Size = $System_Drawing_Size
$buttonUserToAdd.Text = ">>> Add to Role <<<"
$buttonUserToAdd.UseVisualStyleBackColor = $True
$buttonUserToAdd.add_Click($AddUsertoRoleProcess)

$MainWindow.Controls.Add($buttonUserToAdd)


#"Remove from Role" button
$buttonUserToRemove.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonUserToRemove.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 202
$buttonUserToRemove.Location = $System_Drawing_Point
$buttonUserToRemove.Name = "buttonUserToAdd"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 180
$buttonUserToRemove.Size = $System_Drawing_Size
$buttonUserToRemove.Text = ">>> Remove from Role <<<"
$buttonUserToRemove.UseVisualStyleBackColor = $True
$buttonUserToRemove.add_Click($RemoveUserfromRoleProcess)

$MainWindow.Controls.Add($buttonUserToRemove)


#"Create Role Role" button
$buttonCreateRoleGroup.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonCreateRoleGroup.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 232
$buttonCreateRoleGroup.Location = $System_Drawing_Point
$buttonCreateRoleGroup.Name = "buttonCreateRoleGroup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 180
$buttonCreateRoleGroup.Size = $System_Drawing_Size
$buttonCreateRoleGroup.Text = ">>> Create Role Group <<<"
$buttonCreateRoleGroup.UseVisualStyleBackColor = $True
$buttonCreateRoleGroup.add_Click($CreateRoleGroupProcess)

$MainWindow.Controls.Add($buttonCreateRoleGroup)



#"Get Role Role" button
$buttonGetRoleGroup.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonGetRoleGroup.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 740
$System_Drawing_Point.Y = 262
$buttonGetRoleGroup.Location = $System_Drawing_Point
$buttonGetRoleGroup.Name = "buttonCreateRoleGroup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 230
$buttonGetRoleGroup.Size = $System_Drawing_Size
$buttonGetRoleGroup.Text = ">>> Get Role Group Membership <<<"
$buttonGetRoleGroup.UseVisualStyleBackColor = $True
$buttonGetRoleGroup.add_Click($GetRoleGroupProcess)

$MainWindow.Controls.Add($buttonGetRoleGroup)


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
 
