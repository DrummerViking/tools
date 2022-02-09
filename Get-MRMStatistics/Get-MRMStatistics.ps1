<#
.NOTES
	Name: Get-MRMStatus.ps1
	Authors: Agustin Gallegos & Nelson Riera

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Allow admins to check current MRM Statistics and info for users
.DESCRIPTION
	Allow admins to check current MRM Statistics and info for users.
    App brings Current Retention Policy and Tags.
    Can get current Managed Folder Assistant Cycle Stats for primary and Archive Mailbox
    Button available to issue a "Start-ManagedFolderAssistant" on the account
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
    $labelRetentionMenu = New-Object System.Windows.Forms.Label
    $labMbxAlias = New-Object System.Windows.Forms.Label
    $txtBoxMbxAlias = New-Object System.Windows.Forms.TextBox

    $buttonGetMRMPolicy = New-Object System.Windows.Forms.Button
    $buttonGetMRMTags = New-Object System.Windows.Forms.Button
    $buttonGetMRMStatsMbx = New-Object System.Windows.Forms.Button
    $buttonGetMRMStatsArchMbx = New-Object System.Windows.Forms.Button
    $buttonGet7daysStats = New-Object System.Windows.Forms.Button
    $buttonStartMFA = New-Object System.Windows.Forms.Button
    $buttonGetMFALogs = New-Object System.Windows.Forms.Button

    $dgResults = New-Object System.Windows.Forms.DataGridView
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    #endregion Generated Form Objects

    #region connecting to powershell

    # Testing if we have a live PSSession of type Exchange
    # Choosing if connection is to Office 365 or an Exchange on-premises
    $PremiseForm = New-Object System.Windows.Forms.Form
    $radiobutton1 = New-Object System.Windows.Forms.RadioButton
    $radiobutton2 = New-Object System.Windows.Forms.RadioButton
    $buttonGo = New-Object System.Windows.Forms.Button
    $buttonExit = New-Object System.Windows.Forms.Button

    $PremiseForm.Controls.Add($radiobutton1)
    $PremiseForm.Controls.Add($radiobutton2)
    $PremiseForm.Controls.Add($groupbox1)
    $PremiseForm.ClientSize = New-Object System.Drawing.Size(200, 100)
    $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $PremiseForm.Name = "form1"
    $PremiseForm.Text = "Choose your premises"
    #
    # radiobutton1
    #
    $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton1.Location = New-Object System.Drawing.Point(20, 20)
    $radiobutton1.Name = "radiobutton1"
    $radiobutton1.Size = New-Object System.Drawing.Size(100, 25)
    $radiobutton1.TabStop = $True
    $radiobutton1.Text = "Office 365"
    $radioButton1.Checked = $true
    $radiobutton1.UseVisualStyleBackColor = $True
    #
    # radiobutton2
    #
    $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $radiobutton2.Location = New-Object System.Drawing.Point(20, 50)
    $radiobutton2.Name = "radiobutton2"
    $radiobutton2.Size = New-Object System.Drawing.Size(100, 25)
    $radiobutton2.TabStop = $True
    $radiobutton2.Text = "On-Premises"
    $radioButton2.Checked = $false
    $radiobutton2.UseVisualStyleBackColor = $True

    #"Go" button
    $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
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
            if ($radiobutton1.Checked) {
                $Global:premise = "office365"
            }
            else {
                $Global:premise = "on-premises"
            }
            $PremiseForm.Close()
        })
    $PremiseForm.Controls.Add($buttonGo)

    #"Exit" button
    $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonExit.Location = New-Object System.Drawing.Point(120,50)
    $buttonExit.Size = New-Object System.Drawing.Size(50,25)
    $buttonExit.Name = "Exit"
    $buttonExit.Text = "Exit"
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click({ $PremiseForm.Close(); $Global:premise = "exit" })
    $PremiseForm.Controls.Add($buttonExit)

    $InitialMainWindowState = $PremiseForm.WindowState
    $PremiseForm.add_Load($OnLoadMainWindow_StateCorrection)
    $PremiseForm.ShowDialog() | Out-Null

    if ( $Global:premise -eq "exit")
    { return }
    if ( $Global:premise -eq "office365") {
        if ( $null -eq (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue) ) {
            if ($null -eq $cred) { $cred = Get-Credential -Message "Insert your Global Admin credentials" }
            if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
                Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement
            try {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security And Compliance"
                Connect-IPPSSession -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue
            }
            catch {
                if ( ( ($_.Exception.GetBaseException()).errorcode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                    Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Security and Compliance. Requesting to authenticate"
                    Connect-IPPSSession -UserPrincipalName $cred.Username.toString() -ErrorAction Stop -WarningAction SilentlyContinue
                }
                else {
                    return $_
                }
            }
        }
        if ( (Get-PSSession).Computername -notcontains "outlook.office365.com" ) {
            if ($null -eq $cred) { $cred = Get-Credential -Message "Insert your Global Admin credentials" }
            if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
                Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement
            try {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online"
                Connect-ExchangeOnline -Credential $cred -ShowBanner:$False -ErrorAction Stop
            }
            catch {
                if ( ( ($_.Exception.GetBaseException()).errorcode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                    Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Exchange Online. Requesting to authenticate"
                    Connect-ExchangeOnline -UserPrincipalName $cred.Username.toString() -ShowBanner:$False -ErrorAction Stop
                }
                else {
                    return $_
                }
            }
        }
    }
    else {
        # we will test common endpoints for tentative URLs based on
        # autodiscover. domain.com
        # mail .domain.com
        # webmail .domain.com
        $AutoDEmail = . Show-InputBox -Prompt "enter your E-mail Address to discover required Endpoint"
        $AutoDEmail = $AutoDEmail.Substring($AutoDEmail.IndexOf('@') + 1)
        $AutoDEndpoint = $AutoDEmail.Insert(0, "autodiscover.")
        if ($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)) {
            $AutoDEndpoint = $AutoDEmail.Insert(0, "mail.")
            if ($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)) {
                $AutoDEndpoint = $AutoDEmail.Insert(0, "webmail.")
            }
        }
        # if all previous attempts fail, we will request to enter the Exchange Server FQDN or NETBIOS
        if ($null -eq (Test-Connection -ComputerName $AutoDEndpoint -Count 1 -ErrorAction SilentlyContinue)) {
            $AutoDEndpoint = . Show-InputBox -Prompt "Please enter your Exchange CAS FQDN or NETBIOS name"
        }
        # Establishing session
        $Session = New-PSSession -Name Exchange -ConfigurationName Microsoft.Exchange -ConnectionUri http://$AutoDEndpoint/powershell -Authentication Kerberos -AllowRedirection
        Import-PSSession $Session -AllowClobber -CommandName Get-Mailbox, Get-RetentionPolicy, Get-RetentionPolicyTag, Export-MailboxDiagnosticLogs, Start-ManagedFolderAssistant | Out-Null
    }  

    write-host "Warning poped up. Please pay attention" -ForegroundColor white -BackgroundColor Red
    [Microsoft.VisualBasic.Interaction]::MsgBox("This application works when the Primary Mailbox and Archive Mailbox resides in the same premise. 
Unfortunatelly if the Mailbox is on-premises and Archive Online, you can only connect to on-premises and manage on-premises objects.", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
    #endregion

    #region Processes
    #Process to get Current Retention Policy
    $processData = 
    {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList
    
        $policy = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -PropertySets Retention | Select-Object RetentionPolicy 
        if ($null -ne $policy) {
            if ($null -eq $policy.RetentionPolicy) {
                [Microsoft.VisualBasic.Interaction]::MsgBox("User has no Archive mailbox assigned yet or no Default Retention Policy is stamped.", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Warning Message")
            }
            else {
                $array.Add((Get-RetentionPolicy -Identity $policy.RetentionPolicy | Select-Object Name, isDefault, @{Name = "Last Time Modified"; Expression = { $_.WhenChanged } }))
            }
            $dgResults.datasource = $array
            $MainWindow.refresh()

            write-host "Policy displayed" -ForegroundColor white -BackgroundColor Red
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }
        $policy = $null
        $array = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to get associated MRM Policy Tags
    $processData2 = 
    {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList

        $mbx = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -ErrorAction SilentlyContinue -PropertySets Retention | Select-Object Identity, RetentionPolicy
        if ($Null -ne $mbx) {
            $MRMPolicy = Get-RetentionPolicy $mbx.RetentionPolicy -ErrorAction SilentlyContinue
            if ($null -ne $MRMPolicy) {
                foreach ($tag in $MRMPolicy.RetentionPolicyTagLinks) {
                    $taginfo = Get-RetentionPolicyTag $tag | Select-Object Name, Type, messageClass, RetentionAction, AgeLimitForRetention
                    $array.Add($taginfo) | Out-Null
                }
                $dgResults.datasource = $array
                $MainWindow.refresh()
                write-host "Tags displayed" -ForegroundColor white -BackgroundColor Red
            }
            else {
                [Microsoft.VisualBasic.Interaction]::MsgBox("User has no Archive mailbox assigned yet or no Default Retention Policy is stamped.", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Warning Message")
            }
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }

        $mbx = $null
        $MRMPolicy = $null
        $array = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to get current MRM Statistics on Primary mailbox
    $processData3 = 
    {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList

        $mbx = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -ErrorAction Silentlycontinue -PropertySets StatisticsSeed | Select-Object ExchangeGuid
        if ($null -ne $mbx) {
            $Guid = $mbx.ExchangeGuid.Guid
            $logProps = Export-MailboxDiagnosticLogs $Guid -ExtendedProperties
            $xmlprops = [xml]($logProps.MailboxLog)
            $output = $xmlprops.Properties.MailboxTable.Property | Where-Object { $_.name -like "ELC*" } | Select-Object Name, Value
            if ($output) {
                $array.addrange($output)
                $dgResults.datasource = $array
                $MainWindow.refresh()
            }
            write-host "Mailbox Stats displayed" -ForegroundColor white -BackgroundColor Red
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }
        $mbx = $null
        $Guid = $null
        $array = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to get current MRM Statistics on Archive mailbox
    $processData4 = 
    {
        $statusBar.Text = "Running..."
        $array = New-Object System.Collections.ArrayList

        $mbx = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -Archive -ErrorAction SilentlyContinue -PropertySets Archive | Select-Object ArchiveGuid
        if ($null -ne $mbx) {
            $Guid = $mbx.ArchiveGuid.Guid
            $logProps = Export-MailboxDiagnosticLogs $Guid -ExtendedProperties
            $xmlprops = [xml]($logProps.MailboxLog)
            $output = $xmlprops.Properties.MailboxTable.Property | Where-Object { $_.name -like "ELC*" } | Select-Object Name, Value
            if ($output) {
                $array.addrange($output)
                $dgResults.datasource = $array
                $MainWindow.refresh()
            }
            write-host "Archive Stats displayed" -ForegroundColor white -BackgroundColor Red
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }
        $mbx = $null
        $Guid = $null
        $array = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to get last 7 days growth stats
    $processData5 =
    {
        $statusBar.Text = "Running..."
        $user = $txtBoxMbxAlias.Text
        Write-Host "querying user $user content for the last 7 days to determine each day growth" -ForegroundColor Yellow
        $array = New-Object System.Collections.ArrayList

        1..7 | ForEach-Object {
            $startDate = (get-date).AddDays(-$_).tostring("MM-dd-yyyy")
            $endDate = (get-date).AddDays( (-$_ + 1) ).toString("MM-dd-yyyy")   
            $try = Get-ComplianceSearch "$user search$_" -ErrorAction silentlycontinue
            if ( $null -eq $try) {
                Write-Host "querying user $user content from $startDate to $endDate" -ForegroundColor Green
                # Search-Mailbox $user -SearchQuery "Received: $startDate..$endDate" -EstimateResultOnly -DoNotIncludeArchive -SearchDumpster:$False
                $null = New-ComplianceSearch -Name "$user search$_" -ExchangeLocation $user -ContentMatchQuery "Received:$startDate..$endDate"
            }
            else {
                Write-Host "Existing query found for user $user from $startDate to $endDate" -ForegroundColor Green
            }
            Start-ComplianceSearch -Identity "$user search$_" -Force
        }
        # Sleeping 60 seconds to allow searches to complete
        Start-Sleep -Seconds 60

        $i = 0
        1..7 | ForEach-Object {
            $i++
            $startDate = (get-date).AddDays(-$_).tostring("MM-dd-yyyy")
            $endDate = (get-date).AddDays( (-$_ + 1) ).toString("MM-dd-yyyy")
            Write-Host "Getting search results for user $user content from $startDate to $endDate" -ForegroundColor Yellow
             
            $result = ((Get-ComplianceSearch "$user search$_").SearchStatistics | ConvertFrom-Json).ExchangeBinding.Search | Select-Object `
            @{N = "Name"; E = { "$user search$i" } }, `
            @{N = "Search Date Range"; E = { "$startDate to $endDate" } }, `
                ContentItems, ContentSize, HasFaults
            $array.add($result)
        }
        $dgResults.datasource = $array
        $MainWindow.refresh()
        write-host "Last 7 days growth check finished" -ForegroundColor white -BackgroundColor Red
        $array = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to Start Managed Folder Assistant on mailbox
    $processData6 = 
    {
        $statusBar.Text = "Running..."
        $mbx = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -ErrorAction SilentlyContinue
        if ($null -ne $mbx) {
            Start-ManagedFolderAssistant $txtBoxMbxAlias.Text -Verbose
            [Microsoft.VisualBasic.Interaction]::MsgBox("Started successfully", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Managed Folder Assistant")  
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }
        $mbx = $null
        $statusBar.Text = "Process Completed"
    }

    #Process to Get Managed Folder Assistant logs
    $processData7 = 
    {
        $statusBar.Text = "Running..."
        $mbx = Get-EXOMailbox -Identity $txtBoxMbxAlias.Text -ErrorAction SilentlyContinue
        if ($null -ne $mbx) {
            (Export-MailboxDiagnosticLogs $txtBoxMbxAlias.Text -ComponentName mrm -Verbose).Mailboxlog >> $home\Desktop\MFAlog.log
            [Microsoft.VisualBasic.Interaction]::MsgBox("MFA log exported successfully to your desktop", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Managed Folder Assistant")  
            Start-Process notepad.exe -ArgumentList "$home\Desktop\MFAlog.log"
        }
        else {
            write-host "Mailbox not found. Please re type it" -ForegroundColor white -BackgroundColor Red
            [Microsoft.VisualBasic.Interaction]::MsgBox("Mailbox not found", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Mailbox not found")
        }
        $mbx = $null
        $statusBar.Text = "Process Completed"
    }
    #endregion


    $OnLoadMainWindow_StateCorrection =
    { #Correct the initial state of the form to prevent the .Net maximized form issue
        $MainWindow.WindowState = $InitialMainWindowState
    }

    #----------------------------------------------
    #region Generated Form Code
    #main window
    $MainWindow.ClientSize = New-Object System.Drawing.Size(1150,500)
    $MainWindow.DataBindings.DefaultDataSourceUpdateMode = 0
    $MainWindow.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $MainWindow.Name = "Window App"
    $MainWindow.Text = "Managing Retention Policies"
    $MainWindow.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") { $MainWindow.Close() }
    })
    $MainWindow.Icon = $Icon
    $MainWindow.add_Load($handler_MainWindow_Load)
    #
    #dataGrid
    #
    $dgResults.Anchor = 15
    $dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
    $dgResults.DataMember = ""
    $dgResults.Location = New-Object System.Drawing.Point(3,100)
    $dgResults.Size = New-Object System.Drawing.Size(1145,500)
    $dgResults.Name = "dgResults"
    $dgResults.ReadOnly = $True
    $dgResults.RowHeadersVisible = $false
    $dgResults.AllowUserToOrderColumns = $True
    $dgResults.AllowUserToResizeColumns = $True
    $MainWindow.Controls.Add($dgResults)
    #
    #Label "Retention Policies Options" title
    #
    $Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $labelRetentionMenu.Location = New-Object System.Drawing.Point(3,5)
    $labelRetentionMenu.Size = New-Object System.Drawing.Size(200,2)
    $labelRetentionMenu.Name = "Header1"
    $labelRetentionMenu.Text = "Retention Policies Options"
    $labelRetentionMenu.Font = $Font
    $MainWindow.Controls.Add($labelRetentionMenu)
    #
    #Label Mailbox Owner
    #
    $labMbxAlias.Location = New-Object System.Drawing.Point(3,30)
    $labMbxAlias.Size = New-Object System.Drawing.Size(87,20)
    $labMbxAlias.Name = "Mailbox"
    $labMbxAlias.Text = "Check Mailbox"
    $MainWindow.Controls.Add($labMbxAlias)
    #
    #TextBox mailbox Owner
    #
    $txtBoxMbxAlias.DataBindings.DefaultDataSourceUpdateMode = 0
    $txtBoxMbxAlias.Location = New-Object System.Drawing.Point(90,28)
    $txtBoxMbxAlias.Size = New-Object System.Drawing.Size(150,20)
    $txtBoxMbxAlias.Name = "txtBoxMbxAlias"
    #By Default we will populate the user's name running the powershell
    $txtBoxMbxAlias.Text = Show-InputBox -Prompt "Enter the user alias you want to check"
    $MainWindow.Controls.Add($txtBoxMbxAlias)
    #
    # "Get Retention Policy" button
    #
    $buttonGetMRMPolicy.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGetMRMPolicy.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGetMRMPolicy.Location = New-Object System.Drawing.Point(3,57)
    $buttonGetMRMPolicy.Size = New-Object System.Drawing.Size(150,25)
    $buttonGetMRMPolicy.Name = "buttonGetMRMPolicy"
    $buttonGetMRMPolicy.Text = "Get Retention Policy"
    $buttonGetMRMPolicy.UseVisualStyleBackColor = $True
    $buttonGetMRMPolicy.add_Click($processData)
    $MainWindow.Controls.Add($buttonGetMRMPolicy)
    #
    # "Get Retention Policy Tags" button
    #
    $buttonGetMRMTags.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGetMRMTags.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGetMRMTags.Location = New-Object System.Drawing.Point(178,57)
    $buttonGetMRMTags.Size = New-Object System.Drawing.Size(150,25)
    $buttonGetMRMTags.Name = "buttonGetMRMTags"
    $buttonGetMRMTags.Text = "Get Retention Policy Tags"
    $buttonGetMRMTags.UseVisualStyleBackColor = $True
    $buttonGetMRMTags.add_Click($processData2)
    $MainWindow.Controls.Add($buttonGetMRMTags)
    #
    # "Get Mailbox Statistics" button
    #
    $buttonGetMRMStatsMbx.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGetMRMStatsMbx.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGetMRMStatsMbx.Location = New-Object System.Drawing.Point(356,57)
    $buttonGetMRMStatsMbx.Size = New-Object System.Drawing.Size(180,25)
    $buttonGetMRMStatsMbx.Name = "buttonGetMRMStatsMbx"
    $buttonGetMRMStatsMbx.Text = "Get Retention Mailbox Stats"
    $buttonGetMRMStatsMbx.UseVisualStyleBackColor = $True
    $buttonGetMRMStatsMbx.add_Click($processData3)
    $MainWindow.Controls.Add($buttonGetMRMStatsMbx)
    #
    #"Get Archive Mailbox Statistics" button
    #
    $buttonGetMRMStatsArchMbx.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGetMRMStatsArchMbx.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGetMRMStatsArchMbx.Location = New-Object System.Drawing.Point(561,57)
    $buttonGetMRMStatsArchMbx.Size = New-Object System.Drawing.Size(210,25)
    $buttonGetMRMStatsArchMbx.Name = "buttonGetMRMStatsArchMbx"
    $buttonGetMRMStatsArchMbx.Text = "Get Retention Archive Mailbox Stats"
    $buttonGetMRMStatsArchMbx.UseVisualStyleBackColor = $True
    $buttonGetMRMStatsArchMbx.add_Click($processData4)
    $MainWindow.Controls.Add($buttonGetMRMStatsArchMbx)
    #
    # "Get 7 days stats" button
    #
    $buttonGet7daysStats.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGet7daysStats.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGet7daysStats.Location = New-Object System.Drawing.Point(791, 57)
    $buttonGet7daysStats.Size = New-Object System.Drawing.Size(100, 25)
    $buttonGet7daysStats.Name = "7daysStats"
    $buttonGet7daysStats.Text = "Get 7 days stats"
    $buttonGet7daysStats.UseVisualStyleBackColor = $True
    $buttonGet7daysStats.add_Click($processData5)
    $MainWindow.Controls.Add($buttonGet7daysStats)
    #
    # "Start MFA" button
    #
    $buttonStartMFA.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonStartMFA.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonStartMFA.Location = New-Object System.Drawing.Point(910,57)
    $buttonStartMFA.Size = New-Object System.Drawing.Size(100,25)
    $buttonStartMFA.Name = "buttonStartMFA"
    $buttonStartMFA.Text = "Start MFA"
    $buttonStartMFA.UseVisualStyleBackColor = $True
    $buttonStartMFA.add_Click($processData6)
    $MainWindow.Controls.Add($buttonStartMFA)
    #
    #"Get MFA Logs" button
    #
    $buttonGetMFALogs.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonGetMFALogs.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $buttonGetMFALogs.Location = New-Object System.Drawing.Point(1021,57)
    $buttonGetMFALogs.Size = New-Object System.Drawing.Size(100,25)
    $buttonGetMFALogs.Name = "buttonGetMFALogs"
    $buttonGetMFALogs.Text = "Get MFA Log"
    $buttonGetMFALogs.UseVisualStyleBackColor = $True
    $buttonGetMFALogs.add_Click($processData7)
    $MainWindow.Controls.Add($buttonGetMFALogs)
    #endregion Generated Form Code

    #Save the initial state of the form
    $InitialMainWindowState = $MainWindow.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $MainWindow.add_Load($OnLoadMainWindow_StateCorrection)
    #Show the Form
    $MainWindow.ShowDialog() | Out-Null
    if ($MainForm.IsDisposed) {
        Write-Host "Removing temporary ComplianceSearches if any" -ForegroundColor Yellow
        1..7 | ForEach-Object {
            Remove-ComplianceSearch "$user search$_" -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
} #End Function

#Call the Function
GenerateForm