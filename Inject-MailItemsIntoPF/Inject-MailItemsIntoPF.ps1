<#
    .SYNOPSIS
    Script to inject sample messages into a Public Folder.
    
    .DESCRIPTION
    Script to inject sample messages into a Public Folder.
    It requires a Target Public Folder path, and will be validated.
    You can pass a sample file, or if ommitted we will create a test file of 34MB.
    The script will attempt to inject the amount of messages defined in NumberOfMessages.
    Optionally can use BasicAuth (by default will attempt Modern Auth) and enable Transcript.
    
    .PARAMETER TargetPublicFolder
    This is the path to the public folder. It must start with a backslash. For example "\Company Root\folder1\SubFolder2". It should not end with a backslash neither. This is a mandatory parameter.
    
    .PARAMETER SampleFileName
    File path to a sample file to be attach. If this parameter is ommitted, a test file of 34MB will be created.

    .PARAMETER NumberOfMessages
    This is the amount of messages to be created in the Public Folder. By default will attempt 100.
    
    .PARAMETER EnableTranscript
    Use this Switch parameter to enable Powershell Transcript.
    
    .PARAMETER UseBasicAuth
    Use this Switch parameter to connect to EWS using Basic Auth. By default the script will attempt to connect using Modern Auth.
    
    .EXAMPLE
    PS C:\> .\Inject-MailItemsIntoPF.ps1 -TargetPublicFolder "\My Company Root\folder1" -NumberOfMessages 10

    The script will request the user credentials.
    Will validate the folder path exists.
    Will attempt to inject 10 messages into the target folder "\My Company Root\folder1".

    .EXAMPLE
    PS C:\> .\Inject-MailItemsIntoPF.ps1 -TargetPublicFolder "\Corp\subfolder2" -EnableTranscript -UseBasicAuth

    The script will request the user credentials.
    Will validate the folder path exists.
    Will attempt to inject 100 messages (default value) into the target folder.
    Will save all powershell output to Transcript file.
#>
[CmdletBinding()]
param (
    [parameter(Mandatory = $true)]
    [ValidateScript({ # This returns True or False. If 'True', we accept the parameter value. If 'False', we through an error on that value.
            if ( $_.StartsWith("\") -and -not($_.EndsWith("\")) ) { $True }
            else { throw "'$_' is not following the correct syntax. It should start with a backslash, not end with a backslash." }
        })]
    [String] $TargetPublicFolder,

    [String] $SampleFileName = "TestFile.txt",

    [int] $NumberOfMessages = 100,

    [Switch] $EnableTranscript,

    [Switch] $UseBasicAuth
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
    
    # create sample attachment file
    if ( -not(test-path $SampleFileName) ) {
        $file = New-Object -TypeName System.IO.FileStream -ArgumentList "$env:Temp\$SampleFileName", Create, ReadWrite
        $file.SetLength(34MB)
        $file.Close()
    }

    if ($EnableTranscript) {
        Start-Transcript
    }

    $EWS = "$pwd\Microsoft.Exchange.WebServices.dll"
    $test = Test-Path -Path $EWS
    if ($test -eq $False) {
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL in local path not found" -ForegroundColor Cyan
        $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        if ( $null -eq $ewspkg ) {
            Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Downloading EWS DLL Nuget package and installing it" -ForegroundColor Cyan
            $null = Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Trusted -Force
            $null = Install-Package Microsoft.Exchange.WebServices -requiredVersion 2.2.0 -Scope CurrentUser
            $ewspkg = Get-Package Microsoft.Exchange.WebServices -ErrorAction SilentlyContinue
        }        
        $EWSPath = $ewspkg.Source.Replace("\Microsoft.Exchange.WebServices.2.2.nupkg", "")
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in package folder path" -ForegroundColor Cyan
        $EWS = "$EWSPath\lib\40\Microsoft.Exchange.WebServices.dll"
    }
    else {
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] EWS DLL found in current folder path" -ForegroundColor Cyan
    }
    Add-Type -Path $EWS
}

Process {
    # Creating the EWS object
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
    if ( $UseBasicAuth ) {
        [PSCredential]$cred = (Get-credential)
        $exchangeCredentials = New-Object System.Net.NetworkCredential($cred.UserName.ToString(), $cred.GetNetworkCredential().password.ToString())
    }
    else {
        #region Getting oauth credentials using MSAL
        if ( -not(Get-Module MSAL.PS -ListAvailable) ) {
            Install-Module MSAL.PS -Force -ErrorAction Stop
        }
        Import-Module MSAL.PS
        
        # Connecting using Oauth with delegated permissions              
        $ClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db"
        $RedirectUri = "http://localhost/code"
        $scopes = New-Object System.Collections.Generic.List[string]
        $scopes.Add("https://outlook.office365.com/.default")
        #$scopes.Add("https://outlook.office.com/EWS.AccessAsUser.All")
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
        #endregion
    }
    $Service.Credentials = $exchangeCredentials
    $service.EnableScpLookup = $False
    $service.Url = [system.URI]"https://outlook.office365.com/ews/exchange.asmx"
    $service.ReturnClientRequestId = $true
    $service.UserAgent = "InjectMailItemsIntoPF/1.00"

    #region declare findfolder function
    Function Find-Subfolders {
        Param (
            $ParentFolderId,
            $Displayname
        )
        $sourceFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($ParentFolderId)
        $rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$sourceFolderId)
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
        
        $filters = @()
        if ( $null -ne $Displayname){
            $filters += (New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.FolderSchema]::Displayname,$Displayname,[Microsoft.Exchange.WebServices.Data.ContainmentMode]::ExactPhrase, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase))
        }
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND,$filters)
        
        $rootfolder.load()
        foreach ($folder in $rootfolder.FindFolders( $searchFilter, $FolderView ) )
        {   
            $global:bindFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folder.id)
            if ( $bindFolder.DisplayName -eq $Displayname ) {
                $bindFolder.id
                break
            }
        }
    }
    #endregion

    # Binding to the Public Folder Root
    $PFrootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
    $desiredFolderID = $PFrootfolder.id
    
    # looping through each PF level finding each folder in the path, and fetching folderID
    1..(($TargetPublicFolder.Split("\").Count) - 1) | ForEach-Object {
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Looking for folder '$($TargetPublicFolder.Split("\")[$_])'..." -NoNewline
        $desiredFolderID = Find-Subfolders -ParentFolderId $desiredFolderID -Displayname $TargetPublicFolder.Split("\")[$_]
        if ( $null -ne $desiredFolderID ) {
            Write-host "Found" -ForegroundColor Green
        }
    }
    
    # if folder is not found and ID cannot be collected, we will exit the script
    if ( $null -eq $desiredFolderID ) {
        Write-host "Not Found" -ForegroundColor Red
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Target Public Folder '$TargetPublicFolder' not found in the hierarchy. Please check and try again." -ForegroundColor Red
        exit
    }

    # Looping and creating messages in the target PF
    1..$NumberOfMessages | ForEach-Object {
        $subject = "message $_"
        $body = "test message #$_ injected in PF Mailbox"
        Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Creating Email Message object with subject '$Subject'"
        $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
        $Message.Subject = $Subject
        $Message.Body = $Body
        $null = $Message.ToRecipients.Add($token.Result.Account.Username)
        $null = $message.Attachments.AddFileAttachment("$env:Temp\$SampleFileName")
        try {
            $Message.Save($desiredFolderID)
        }
        catch {
            if ( $_.Exception.Message.contains("The server cannot service this request right now. Try again later") ) {
                Write-host "[$((Get-Date).ToString("HH:mm:ss"))] Throttling occurred. Skipping message and continue." -ForegroundColor Red
            }
        }
    }
}

End {
    if ( $EnableTranscript) {
        Stop-Transcript
    }
}