[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True, HelpMessage = 'Primary Email Address of a cloud mailbox you want to check...')]
    [string]$CloudUser = ''
)

# Disclaimer
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
# Installing PSFramework
if ( !(Get-Module PSFramework) -and !(Get-Module PSFramework -ListAvailable) )
{
    Install-Module PSFramework -Force
}

#connecting to Cloud side
if ( -not (Get-PSSession | Where-Object computername -eq "outlook.office365.com") ) {
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online Powershell"

    $LiveCred = Get-Credential -Message "Please enter your Global Admin Credentials"
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) )
    {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -Credential $LiveCred
}
$ts = Get-Date -Format "yyyy-MM-dd hh_mm_ss"
$FormatEnumerationLimit = -1

<#using C:\TEMP\MSlogs folder
Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Checking C:\TEMP\MSlogs Folder"
$folder = "C:\TEMP\MSlogs" 
if (-not (Test-path $folder) )
{ 
    #Create the directory 
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Creating Directory $folder" -DefaultColor Green 
    $null = [System.IO.Directory]::CreateDirectory($folder) 
} 
else
{ 
    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] $folder is already created!" -DefaultColor Yellow 
}
Set-Location C:\temp\MSlogs
#>

#setting variables
if($CloudUser -eq '') { $CloudUser = Read-Host -Prompt "please enter the Primary Email Address of a cloud mailbox" }

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting OrganizationConfig settings"
$orgConfig = Get-OrganizationConfig

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting TransportConfig settings"
$transportConfig = Get-TransportConfig

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting User's CasMailbox settings"
$CasMbx = Get-EXOCasMailbox $CloudUser -Properties SmtpClientAuthenticationDisabled

Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Collecting Exchange's Authentication Policies settings"
$AuthPolicies = Get-AuthenticationPolicy

[PSCustomObject]$Data = @{
Mailbox = $CloudUser
SMTPLegacyAuthMailboxLevel = -not $CasMbx.SmtpClientAuthenticationDisabled
ModernAuthEnabled = $orgConfig.OAuth2ClientProfileEnabled
SMTPLegacyAuthOrganizationLevel = -not $transportConfig.SmtpClientAuthenticationDisabled
}
$Data | Format-Table -AutoSize