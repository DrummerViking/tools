﻿[CmdletBinding()]
param(
    [switch] $EnableTranscript = $false,

    [String] $CSVFilePath,

    [switch] $ValidateRoomsExistence
)

Begin {
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
    if ($EnableTranscript) {
        Start-Transcript
    }
    $EWS = "$pwd\Microsoft.Exchange.WebServices.dll"
    $test = Test-Path -Path $EWS
    if ($test -eq $False) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL in local path not found" -ForegroundColor Cyan
        $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        if ( $null -eq $ewspkg ) {
            Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Downloading EWS DLL Nuget package and installing it" -ForegroundColor Cyan
            $null = Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -Force
            $null = Install-Package Microsoft.Exchange.WebServices -requiredVersion 2.2.0 -Scope CurrentUser
            $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        }        
        $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg","")
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in package folder path" -ForegroundColor Cyan
        $EWS = "$EWSPath\lib\40\Microsoft.Exchange.WebServices.dll"
    }
    else {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in current folder path" -ForegroundColor Cyan
    }
    Add-Type -Path $EWS
}

process {
    #region Selecting CSV file
    if ( $CSVFilePath.Length -eq 0 ) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Please pick up the CSV files with the list of previous and new rooms to replace."
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        [System.Windows.Forms.Application]::EnableVisualStyles() 
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $PSScriptRoot
        $OpenFileDialog.ShowDialog() | Out-Null
        if ($OpenFileDialog.filename -ne "") {
            $CSVPath = $OpenFileDialog.filename
        }
    }else {
        $global:CSVPath = $CSVFilePath
    }
    $csv = import-csv $CSVPath
    if ( ($csv | Get-Member -Name PreviousRoom,NewRoom).count -lt 2 ){
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] CSV file does not contain the necessary columns. Please check you have 'PreviousRoom, NewRoom' columns and try again." -ForegroundColor Red
        return
    }
    #endregion

    #region Validate if room mailboxes exists as valid recipients
    if ( $ValidateRoomsExistence ) {
        if ( (Get-PSSession).Computername -notcontains "outlook.office365.com" ) {
            if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable) ) {
                Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement
            Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online"
            Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
        }
        foreach ($line in $csv) {
            try {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Checking if Room mailbox $($line.NewRoom) exists..." -NoNewline
                $null = Get-EXORecipient $line.newRoom -ErrorAction Stop
                Write-host "Ok." -ForegroundColor Green

                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Checking if Room mailbox $($line.PreviousRoom) exists..." -NoNewline
                $null = Get-EXORecipient $line.PreviousRoom -ErrorAction Stop
                Write-host "Ok." -ForegroundColor Green
            }
            catch {
                Write-host "Failed." -ForegroundColor Red
                $failedAlias = $_.Exception.Message.split("'")[2]
                Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Room mailbox $failedAlias not found. Exiting script." -ForegroundColor Red
                exit
            }
        }      
    }

    #creating service object
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
    
    #region Getting oauth credentials using MSAL
    if ( -not(Get-Module Microsoft.Identity.Client -ListAvailable) ) {
        Install-Module Microsoft.Identity.Client -Force -ErrorAction Stop
    }
    Import-Module Microsoft.Identity.Client
    $AppId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
    $pcaOptions = [Microsoft.Identity.Client.PublicClientApplicationOptions]::new()
    $pcaOptions.ClientId = $AppId
    $pcaOptions.RedirectUri = "http://localhost/code"
    $pcaBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($pcaOptions)
    $pca = $pcaBuilder.Build()
    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    #$scopes.Add("https://outlook.office.com/EWS.AccessAsUser.All")
    $authResult = $pca.AcquireTokenInteractive($scopes)
    $global:token = $authResult.WithPrompt()
    if ($Token.Status -eq "faulted") {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Failed to obtain authentication token. Exiting script." -ForegroundColor Red
        exit
    }
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token.Result.AccessToken)
    $service.Url = New-Object Uri("https://outlook.office365.com/ews/exchange.asmx")
    $Service.Credentials = $exchangeCredentials
    #endregion
}

End {
    if ($EnableTranscript) {
        Stop-Transcript
    }
}