<#
.NOTES
	Name: Parse incident Report emails.ps1
	Author: Agustin Gallegos
 
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.SYNOPSIS
    Export all Incident Report emails in one folder to CSV or XLSX
.PARAMETER Recipients
    The email address you want the report to be sent to.
.PARAMETER OrgAdmins
    Send report to Organization Admins detected.
.COMPONENT
   DLP
.ROLE
   Support
#>
Param(
    [Parameter(Position = 1, Mandatory = $False, HelpMessage = 'The email address you want the report to be sent to...')]
    [string[]]$recipients = '',
    [Parameter(Position = 2, Mandatory = $False, HelpMessage = 'Send report to Organization Admins detected...')]
    [Switch]$OrgAdmins = $False		
)

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

#region import EWS DLL file# Locating DLL location either in working path, in EWS API 2.1 path or in EWS API 2.2 path
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
    $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg","")
    Write-Host "EWS DLL found in package folder path" -ForegroundColor Cyan
    $EWS = "$EWSPath\lib\40\Microsoft.Exchange.WebServices.dll"
}
else {
    Write-Host "EWS DLL found in current folder path" -ForegroundColor Cyan
}
Add-Type -Path $EWS
#endregion

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

#Getting oauth credentials
if ( -not(Get-Module MSAL.PS -ListAvailable) ) {
    Install-Module MSAL.PS -Force -ErrorAction Stop
}
Import-Module MSAL.PS

# Connecting using Oauth with delegated permissions              
$ClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
$RedirectUri = "http://localhost/code"
$scopes = New-Object System.Collections.Generic.List[string]
$scopes.Add("https://outlook.office365.com/.default")
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
$exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.AccessToken)
$service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
$Service.Credentials = $exchangeCredentials
$service.ReturnClientRequestId = $true
$service.UserAgent = "ParseIncidentReports/1.80"

Function Find-Subfolders {
    Param (
        $ParentFolderId,

        $ParentDisplayname
    )
    $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($ParentFolderId)
    $service.ClientRequestId = (New-Guid).ToString()
    $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)

    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    #$FolderView.Traversal = "Deep"
    
    $service.ClientRequestId = (New-Guid).ToString()
    $rootfolder.load()
    $service.ClientRequestId = (New-Guid).ToString()
    foreach ($folder in $rootfolder.FindFolders($FolderView) ) {
        $DisplayName = "$ParentDisplayname\$($Folder.Displayname)"
        $output = $folder | Select-Object @{N = "Displayname" ; E = {$DisplayName}}, Id
        if ($folder.ChildFolderCount -gt 0) {
            #write-host "looking for subfolders under $($folder.displayname)" -ForegroundColor Green
            $service.ClientRequestId = (New-Guid).ToString()
            Find-Subfolders -ParentFolderId $folder.id -ParentDisplayname $Displayname -Array $array
        }
        $output
    }
}
$rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
$folders = Find-Subfolders -ParentFolderId $rootfolder.id -ParentDisplayname ""
[int]$index = 0
$folders | ForEach-Object {
    [PSCustomObject] @{ Order = $index; DisplayName = $_.DisplayName }; $index++
} | Out-Host
Do {
    Write-Host "Choose the order number for the folder where your Incident reports are located: " -NoNewline
    [int]$option = Read-Host
} while ( $option -lt 0 -or $option -ge $folders.Count)

$filename = "$env:userprofile\desktop\Incident Report - $((Get-Date).ToString("MM-dd-yyyy HH_mm_ss")).csv"
"Received Time, Report Id, Message Id, sender, Subject, To, Rule Hit" | Out-File $filename -Encoding ascii -Append
Write-Host "Report File generated: $filename" -foregroundColor Green
Write-Host "Please wait while we work on the incidents reports" -foregroundColor Green

$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folders[$option].Id)

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(250)
$nl = "`n`r"
$fiItems = $null
do {  
    $fiItems = $service.FindItems($Folder.Id, $ivItemView)  
    foreach ($Item in $fiItems.Items) {  
        $TempItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service, $Item.Id)
        $text = $TempItem.Body.Text.Replace("<br>", $nl)
        $text = $text.split($nl)

        $ReportId = $text | Select-String -SimpleMatch "Report Id: " | Out-String
        $MessageID = $text | Select-String -SimpleMatch "Message Id: " | Out-String
        $senderAddress = $text | Select-String -SimpleMatch "Sender: " | Out-String
        $Subject = $text | Select-String -SimpleMatch "Subject: " | Out-String
        $To = $text | Select-String -SimpleMatch "To: " | Out-String
        $RuleHit = $text | Select-String -SimpleMatch "Rule Hit: " | Out-String
                
        $ReportId = $ReportId.TrimEnd()
        $MessageID = $MessageID.TrimEnd()
        $senderAddress = $senderAddress.TrimEnd()
        $Subject = $Subject.TrimEnd()
        $To = $To.TrimStart() ; $To = $To.TrimEnd()
        if ($RuleHit.Contains("$nl")) {
            $RuleHit = $RuleHit.Replace("$nl", "")
            $RuleHit = $RuleHit.TrimEnd()
        }

        $ReportId = $ReportId.Substring($ReportId.IndexOf('Report Id:') + 11)
        $MessageID = $MessageID.Substring($MessageID.IndexOf(':') + 2)
        $MessageID = $MessageID.Replace("&lt;", "<")
        $MessageID = $MessageID.Replace("&gt;", ">")
        $senderAddress = $senderAddress.Substring($senderAddress.IndexOf(':') + 2)
        if ($senderAddress.Contains(",")) {
            $senderAddress = $senderAddress.Split(",")[1]
        }
        $Subject = $Subject.Substring($Subject.IndexOf(':') + 2)
        $to -match "[\w-\.]+@([\w-]+\.)+[\w-]{2,4}" | Out-Null
        $to = $Matches.Values -join ";"
        <#
        $To = $To.Substring($To.IndexOf(':') + 2)
        if ($To.Contains(",")) {
            $To = $To.Split(",")[1]
            $To = $To.TrimStart()
            $To = $To.Split("$nl")[0]
            $To = $To.TrimEnd()
        }
        #>
        $RuleHit = $RuleHit.Replace("Rule Hit: ","")
        $RuleHit = $RuleHit.Replace(",", ";")
                                                
        $out = $TempItem.DateTimeReceived.GetDateTimeFormats()[14].toString() + "," + $ReportId + "," + $MessageID + "," + $senderAddress + "," + $Subject + "," + $To + "," + $RuleHit     
        $out | Out-File $filename -Append -Encoding ascii
    }  
    $ivItemView.Offset += $fiItems.Items.Count  
}while ($fiItems.MoreAvailable -eq $true)  

Write-Host "Report File finished." -foregroundColor Green

if ($recipients -ne '' -or $OrgAdmins -ne $False) {
    #region parameters
    $listrecipients = New-Object System.Collections.ArrayList

    # If Switch $OrgAdmins is in use, we will check current admins and include them to the recipients list
    if ($OrgAdmins -eq $True) {
        if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable )) {
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        }
        if (-not(Get-ConnectionInformation) ) {
            $script:cred = Get-Credential
            Connect-ExchangeOnline -Credential $script:cred -ShowBanner:$false
        }
        
        $TenantAdmins = Get-RoleGroupMember "Organization Management"
        foreach ($admin in (Get-RoleGroupMember $TenantAdmins.Name)) {
            if ($Recipients -ne '') {
                $Recipients = $Recipients + ", "
            }
            $Recipients = $Recipients + (Get-Mailbox $admin.Name).PrimarySmtpAddress
        }
    }
    $listrecipients = ("$Recipients").Split(",")
    Write-Host "Sending Report by e-mail to:" $recipients -ForegroundColor Yellow
    Write-Host
    #endregion parameters


    # sending message
    $Subject = "Incident Report - $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))"
    Send-MailMessage -From $cred.UserName -To $listrecipients -Body "Incident report generated on $((Get-Date).ToString("MM/dd/yyyy HH:mm:ss"))" -SmtpServer smtp.office365.com -UseSsl -Port 587 -Subject $Subject -Credential $script:cred -Attachments $filename
}