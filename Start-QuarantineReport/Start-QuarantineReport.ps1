﻿<#
.NOTES
	Name: Start-QuarantineReport.ps1
	Authors: Agustin Gallegos
	Version History:
    2.00 - 01/24/2022 - Updated script to Github repository.
                      - Updated script to use new EXO PS module.
                      - Added additional modules dependency.
    1.00 - 03/02/2018 - First Release
	1.00 - 03/02/2018 - Project start    

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 
.SYNOPSIS
    Automatically generate HTML report listing quarantine messages for a Security Group and its members.

.DESCRIPTION
    Automatically generate HTML report listing quarantine messages for a Security Group and its members.
    This report will be send by e-mail to a list of recipients.

.PARAMETER GroupAddress     
    Group Alias you want to get the list of members of

.PARAMETER recipients
    comma separated list of recipients to which the report should be sent to.

.PARAMETER OrgAdmins
    This is a switch Parameter. Using it, will send the report to every Global Admin in the tenant. Can be combined together with "recipients" parameter.

.PARAMETER EmailtoGroupMembers
    This is a switch Parameter. Using it, will send the report to the group you are collecting the report for.

.EXAMPLE 
    .\QuarantinePerGroupReport.ps1 -GroupAlias InfoSecurity -Recipients "agallego@Outlook.com"
 
.EXAMPLE 
    .\QuarantinePerGroupReport.ps1 -GroupAlias HR -OrgAdmins -EmailtoGroupMembers

.COMPONENT
   AntiSpam
.ROLE
   Support
#>

Param(
    [Parameter(Position = 1, Mandatory = $True, HelpMessage = 'The group SMTP address you want to get members of...')]
    [string]$GroupAddress = '',
    [Parameter(Position = 2, Mandatory = $False, HelpMessage = 'The email address you want the report to be sent to...')]
    [string]$Recipients = '',
    [Parameter(Position = 3, Mandatory = $False, HelpMessage = 'send report to Organization Admins detected...')]
    [Switch]$OrgAdmins = $False,
    [Parameter(Position = 4, Mandatory = $False, HelpMessage = 'The group you want to get members of...')]
    [Switch]$EmailToGroupMembers = $False
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

if ($recipients -eq '' -and $OrgAdmins -eq $false -and $EmailtoGroupMembers -eq $False) {
    write-host "You need to select at least 1 recipient, OrgAdmins or EmailtoGroupMembers switch in order to continue" -ForegroundColor White -BackgroundColor Red
    return
}

# Connect to EXO if no existing Session available
if ((Get-PSSession).Computername -notlike "*outlook*") {
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}

Write-Host "Getting list of Group members and formatting to HTML"
Write-Host
Get-DistributionGroupMember $GroupAddress | Select-Object PrimarySMTPAddress | Export-Csv -Path $PSScriptRoot\mailboxes.csv -NoTypeInformation

Write-Host "Creating temporary file to list all members of the group:" $PSScriptRoot\mailboxes.csv -ForegroundColor Yellow
$mbxs = Import-Csv -Path $PSScriptRoot\mailboxes.csv

# creating variable to store the information we will later output to HTML
$DisplayList = @()

# Looping through each mailbox
foreach ($mbx in $mbxs.PrimarySMTPAddress) {
    # creating variable to store user data
    $DisplayList += $DisplayList = Get-QuarantineMessage -RecipientAddress $mbx | Select-Object ReceivedTime, SenderAddress, @{N = "RecipientAddress"; E = { $mbx } }, Subject, Size, Type, QuarantineTypes, Expires, Direction
    
}
[string]$html = $DisplayList | ConvertTo-Html

Write-Host "Removing temporary file:" $PSScriptRoot\mailboxes.csv -ForegroundColor Yellow
Remove-Item $PSScriptRoot\mailboxes.csv -Force -Confirm:$false

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
			')
$HTML = $HTML.Replace('</tr> <tr>', '</tr> <tr style=''background-color:#BBD9EE''>')

#region parameters
$listrecipients = New-Object System.Collections.ArrayList

#if EmailtoGroupMembers is in use, we will add group's email address to the recipients list
if ($EmailtoGroupMembers -eq $True) {
    $Recipients = $Recipients + ", " + $GroupAddress
}
# If Switch $OrgAdmins is in use, we will check current admins and include them to the recipients list
if ($OrgAdmins -eq $True) {
    $TenantAdmins = Get-RoleGroupMember "Organization Management"
    foreach ($admin in (Get-RoleGroupMember $TenantAdmins.Name)) {
        if ($Recipients -ne '') {
            $Recipients = $Recipients + ", "
        }
        $Recipients = $Recipients + (Get-Mailbox $admin.Name).PrimarySmtpAddress
    }
}
$listrecipients = ("$Recipients").Split(",")
#endregion parameters

# generating Subject
$Subject = "Quarantine Group Report $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))"

# Saving report in desktop
Write-Host "Saving report to $Home\Desktop\Quarantine Report.html"
$html | Add-Content -Path "$Home\Desktop\Quarantine Report.html" -Force

# sending message
Write-Host "Sending Report by e-mail to" $recipients
Write-Host
if ($Null -eq $cred) {
    $cred = Get-Credential -Message "Type your Sender's credentials"
}
Send-MailMessage -From $cred.UserName -To $listrecipients -Body $html -BodyasHtml -SmtpServer smtp.office365.com -UseSsl -Port 587 -Subject $Subject -Credential $cred