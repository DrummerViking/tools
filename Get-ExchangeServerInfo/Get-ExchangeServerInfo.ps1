<#
.NOTES
	Name: Get-ExchangeServerInfo.ps1
	Author: Agustin Gallegos
	Requires: Exchange Management Shell and administrator rights
 
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Checks the Exchange servers for general information
.DESCRIPTION
	This script checks the On-Premises Exchange servers for general information by getting Server's name, build number (and RU/CU name), AD Site and roles, .NET version, V++ versions.
    If you have a mixed environment running multiple versions of Exchange, it is recommended to run the script from your newest version available.
    Includes a "CASLoad" switch, to collect some Protocol load counters, Netlogon's Semaphores, and check ASA credentials if it is running in a local CAS Server.
    Includes an "Autodiscover" switch, to collect SCPs.
.PARAMETER Autodiscover
    This optional parameter provides information about Autodiscover in the selected CAS servers.
    It will provide Server name, RU, Server's Site, SCP Endpoint, and Autodiscover SiteScope.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified. 
    If it is not, it will look up for all servers in the Organization.
.PARAMETER Site
    This optional parameter allows the target Site to be specified, so the script will look up for all servers in the specified Site. 
    If it is not, it will look up for all servers in the Organization.
.PARAMETER CASLoad
    This optional switch allows to collect CAS related information. Servers without the CAS role will be skipped.
    It will check for Netlogon's semaphores timeouts.
    It will check for the most common protocols user load.
.PARAMETER CheckASACredentails
    This optional switch allows to check for ASA Credentails in the environment.
.EXAMPLE
	.\Get-ExchangeServerInfo.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\Get-ExchangeServerInfo.ps1 -Server SERVERNAME -CASLoad
	Run against a single remote Exchange server including CAS related information.
.EXAMPLE
    .\Get-ExchangeServerInfo.ps1 -Site SITENAME
    Run against all servers located in the specified Site.
#>
 
Param(
    $Server = '*',
    $Site = '*',
    [Switch]$CASLoad = $false,
    [Switch]$CheckASACredentials = $false,
    [Switch]$Autodiscover = $false
)
 
    Function Get-ServerVersion{
    param(
        $CAS = ''
    )
    #setting Variable
    $exSetupVer = $null

                if ($CAS.AdminDisplayVersion.Major -eq "15"){
		                $Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v15\\Setup"
	                }
	                elseif ($CAS.AdminDisplayVersion.Major -eq "14")
	                {
		                $Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v14\\Setup"
	                }
	                elseif    ($CAS.AdminDisplayVersion.Major -eq "8")
	                {
		                $Reg_ExSetup = "SOFTWARE\\Microsoft\\Exchange\\Setup"
	                }
                # Read Rollup Update information from servers
                # Set Registry constants
	                $VALUE1 = "MsiInstallPath"
 
                # Open remote registry
	                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $CAS)
 
                # Set regKey for MsiInstallPath
	                $regKey= $reg.OpenSubKey($Reg_ExSetup)
 
                # Get Install Path from Registry and replace : with $
	                $installPath = ($regKey.getvalue($VALUE1) | ForEach-Object {$_ -replace (":","`$")})
                # Set ExSetup.exe path
	                $binFile = "Bin\ExSetup.exe"
                # Get ExSetup.exe file version
                    $exSetupVer = ((Get-Command "\\$CAS\$installPath$binFile").FileVersionInfo | ForEach-Object {$_.FileVersion})
                
                # Get RU/CU version
                    [int]$CU = $exSetupVer.subString(6,4)
                    # Determine RU version for Exchange 2007 SP3 builds
                    if($CAS.AdminDisplayVersion.Major -eq "8"){
                        switch($CU)
                        {
                            0083 {$exSetupVer += " - RTM"}
                            0106 {$exSetupVer += " - RU1"}
                            0137 {$exSetupVer += " - RU2"}
                            0159 {$exSetupVer += " - RU3"}
                            0192 {$exSetupVer += " - RU4"}
                            0213 {$exSetupVer += " - RU5"}
                            0245 {$exSetupVer += " - RU6"}
                            0264 {$exSetupVer += " - RU7"}
                            0279 {$exSetupVer += " - RU8.3"}
                            0297 {$exSetupVer += " - RU9"}
                            0298 {$exSetupVer += " - RU10"}
                            0327 {$exSetupVer += " - RU11"}
                            0342 {$exSetupVer += " - RU12"}
                            0348 {$exSetupVer += " - RU13"}
                            0379 {$exSetupVer += " - RU14"}
                            0389 {$exSetupVer += " - RU15"}
                            0406 {$exSetupVer += " - RU16"}
                            0417 {$exSetupVer += " - RU17"}
                            0445 {$exSetupVer += " - RU18"}
                            0459 {$exSetupVer += " - RU19"}
                            0468 {$exSetupVer += " - RU20"}
                            0485 {$exSetupVer += " - RU21"}
                            0502 {$exSetupVer += " - RU22"}
                            0517 {$exSetupVer += " - RU23"}
                            default {}
                        }
                    }
 
                    # Determine RU version for Exchange 2010 SP3 builds
                    if($CAS.AdminDisplayVersion.Major -eq "14"){
                        switch($CU)
                        {
                            0123 {$exSetupVer += " - RTM"}
                            0146 {$exSetupVer += " - RU1"}
                            0158 {$exSetupVer += " - RU2"}
                            0169 {$exSetupVer += " - RU3"}
                            0174 {$exSetupVer += " - RU4"}
                            0181 {$exSetupVer += " - RU5"}
                            0195 {$exSetupVer += " - RU6"}
                            0210 {$exSetupVer += " - RU7"}
                            0224 {$exSetupVer += " - RU8.2"}
                            0235 {$exSetupVer += " - RU9"}
                            0248 {$exSetupVer += " - RU10"}
                            0266 {$exSetupVer += " - RU11"}
                            0279 {$exSetupVer += " - RU12"}
                            0294 {$exSetupVer += " - RU13"}
                            0301 {$exSetupVer += " - RU14"}
                            0319 {$exSetupVer += " - RU15"}
                            0339 {$exSetupVer += " - RU16"}
                            0351 {$exSetupVer += " - RU17"}
                            0361 {$exSetupVer += " - RU18"}
                            0382 {$exSetupVer += " - RU19"}
                            0389 {$exSetupVer += " - RU20"}
                            0399 {$exSetupVer += " - RU21"}
                            0411 {$exSetupVer += " - RU22"}
                            0417 {$exSetupVer += " - RU23"}
                            0419 {$exSetupVer += " - RU24"}
                            0435 {$exSetupVer += " - RU25"}
                            0442 {$exSetupVer += " - RU26"}
                            0452 {$exSetupVer += " - RU27"}
                            0461 {$exSetupVer += " - RU28"}
                            0468 {$exSetupVer += " - RU29"}
                            0496 {$exSetupVer += " - RU30"}
                            default {}
                        }
                    }
 
                    
                    if($CAS.AdminDisplayVersion.Major -eq "15"){
                        # Determine CU version for Exchange 2013 builds
                        if($CAS.AdminDisplayVersion.Minor -eq "0"){
                            switch($CU)
                            {
                                0516 {$exSetupVer += " - RTM"}
                                0620 {$exSetupVer += " - CU1"}
                                0712 {$exSetupVer += " - CU2"}
                                0775 {$exSetupVer += " - CU3"}
                                0847 {$exSetupVer += " - CU4"}
                                0913 {$exSetupVer += " - CU5"}
                                0995 {$exSetupVer += " - CU6"}
                                1044 {$exSetupVer += " - CU7"}
                                1076 {$exSetupVer += " - CU8"}
                                1104 {$exSetupVer += " - CU9"}
                                1130 {$exSetupVer += " - CU10"}
                                1156 {$exSetupVer += " - CU11"}
                                1178 {$exSetupVer += " - CU12"}
                                1210 {$exSetupVer += " - CU13"}
                                1236 {$exSetupVer += " - CU14"}
                                1263 {$exSetupVer += " - CU15"}
                                1293 {$exSetupVer += " - CU16"}
                                1320 {$exSetupVer += " - CU17"}
                                1347 {$exSetupVer += " - CU18"}
                                1365 {$exSetupVer += " - CU19"}
                                1367 {$exSetupVer += " - CU20"}
                                1395 {$exSetupVer += " - CU21"}
                                1473 {$exSetupVer += " - CU22"}
                                1497 {$exSetupVer += " - CU23"}
                                default {}
                            }
                        # Determine CU version for Exchange 2016 builds
                        }elseif($CAS.AdminDisplayVersion.Minor -eq "1"){
                           switch($CU)
                            {
                                0225 {$exSetupVer += " - RTM"}
                                0396 {$exSetupVer += " - CU1"}
                                0466 {$exSetupVer += " - CU2"}
                                0544 {$exSetupVer += " - CU3"}
                                0669 {$exSetupVer += " - CU4"}
                                0845 {$exSetupVer += " - CU5"}
                                1034 {$exSetupVer += " - CU6"}
                                1261 {$exSetupVer += " - CU7"}
                                1415 {$exSetupVer += " - CU8"}
                                1466 {$exSetupVer += " - CU9"}
                                1531 {$exSetupVer += " - CU10"}
                                1591 {$exSetupVer += " - CU11"}
                                1713 {$exSetupVer += " - CU12"}
                                1779 {$exSetupVer += " - CU13"}
                                1847 {$exSetupVer += " - CU14"}
                                1913 {$exSetupVer += " - CU15"}
                                1979 {$exSetupVer += " - CU16"}
                                2044 {$exSetupVer += " - CU17"}
                                default {} 
                            }
                        }
                        # Determine CU version for Exchange 2019 builds
                        }elseif($CAS.AdminDisplayVersion.Minor -eq "2"){
                            switch($CU)
                            {
                                0221 {$exSetupVer += " - RTM"}
                                0330 {$exSetupVer += " - CU1"}
                                0397 {$exSetupVer += " - CU2"}
                                0464 {$exSetupVer += " - CU3"}
                                0529 {$exSetupVer += " - CU4"}
                                0595 {$exSetupVer += " - CU5"}
                                0659 {$exSetupVer += " - CU6"}
                                default {} 
                            }
                        }
                    
    }


 
# Declaring some variables 
$totalOutlook = 0
$totalOA = 0
$totalMapiHTTP = 0
$totalOWA = 0
$netver = '' 
 
# If the Switch "CheckASACredentials" is in place, we will look up for an ASA credential locally
            if($CheckASACredentials -eq $true){
                $CAS = get-ExchangeServer $env:COMPUTERNAME
                    if($CAS.AdminDisplayVersion.Major -eq "8"){
                    Write-Host "Your local server is running Exchange 2007. You need to run Exchange 2010 or above to validate ASA credentials." -ForegroundColor Yellow
                    }else{
                        if($CAS.isClientAccessServer -eq $true){
                            write-host "Checking if ASA Credential is in place in the environment" -ForegroundColor Yellow
                            Write-Host "it will be checked against the local CAS server $CAS" -ForegroundColor Yellow
                            Write-Host "take into account that we assume the same credentials are in place for all CAS servers with the same Exchange version and in the same AD site" -ForegroundColor Yellow
                            
                            $result = (Get-ClientAccessServer $env:COMPUTERNAME -IncludeAlternateServiceAccountCredentialPassword).AlternateServiceAccountConfiguration
                            $result
						Write-Host " "
                        }else{
                            write-host "Checking if ASA Credential is in place in the environment" -ForegroundColor Yellow
                            Write-Host "it will be checked against the local server $CAS" -ForegroundColor Yellow
                            Write-Host "The local machine is not a CAS server. Please re run the script from a CAS server"
                        }
                    }
                }
 
#If "Autodiscover" switch is used, we will display Autodiscover related info
if($Autodiscover -eq $true){
    $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True"}
    # Check if "Server" parameter was in used
    if($Server -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True" -and $_.Name -like "*$Server*"} | Sort-Object Name
        }
        # Check if "Site" parameter was in used
        elseif($Site -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True" -and $_.Site -like "*$Site*"} | Sort-Object Name
        Write-Host "Getting CAS servers from Site: $Site" -ForegroundColor Green
    }
    # Creating Array variable to collect Server's information
    $ServersList = @()
 
    foreach($CAS in $Servers){
            Write-Host "Validating info against CAS Server $CAS" -ForegroundColor Green
            $tempServer = New-Object System.Object
            $tempServer | Add-Member -Type NoteProperty -Name ServerName -Value $CAS
            $tempServer | Add-Member -Type NoteProperty -Name Site -Value $CAS.Site.Rdn.UnescapedName
            $tempServer | Add-Member -Type NoteProperty -Name "SCP Endpoint" -Value (Get-ClientAccessServer $CAS.name -warningAction SilentlyContinue).AutoDiscoverServiceInternalUri.Authority
            $tempServer | Add-Member -Type NoteProperty -Name "Site Scope" -Value (Get-ClientAccessServer $CAS.name -warningAction SilentlyContinue).AutoDiscoverSiteScope
                       
            $ServersList += $tempServer
    }
    
}
 
 
 
#region If "CASLoad" switch is used, we will collect CAS related counters info
if($CASLoad -eq $true -and $Autodiscover -eq $false){
    $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True"}
    # Check if "Server" parameter was in used
    if($Server -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True" -and $_.Name -like "*$Server*"} | Sort-Object Name
        }
        # Check if "Site" parameter was in used
        elseif($Site -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.IsClientAccessServer -eq "True" -and $_.Site -like "*$Site*"} | Sort-Object Name
        Write-Host "Getting CAS servers from Site: $Site" -ForegroundColor Green
    }
    
    # Creating Array variable to collect Server's information
    $ServersList = @()
 
    foreach($CAS in $Servers){
        # Testing if the server is reachable
        $conn = Test-Connection -ComputerName $CAS -Count 1 -ErrorAction SilentlyContinue
 
        # Getting counters only if connection is alive
        if($null -ne $conn){
            Write-Host "Validating info against CAS Server $CAS" -ForegroundColor Green
            $tempServer = New-Object System.Object
            $tempServer | Add-Member -Type NoteProperty -Name ServerName -Value $CAS
            . Get-ServerVersion -CAS $CAS
            $tempServer | Add-Member -Type NoteProperty -Name "ServerBuild          " -Value $exSetupVer
            $tempServer | Add-Member -Type NoteProperty -Name Site -Value $CAS.Site.Rdn.UnescapedName
 
            $tempServer | Add-Member -Type NoteProperty -Name OutlookRpc -Value 0
            # We are excluding Exchange 2007 version, as RPC does not go to a CAS box, but to the user's Mailbox server.
            if($CAS.AdminDisplayVersion -notlike "*8.*"){
                $OutlookCounter = (Get-Counter -counter "\MSExchange RpcClientAccess\User Count" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
                $tempServer.OutlookRpc = $OutlookCounter
                $totalOutlook += $OutlookCounter
            }
        
            $tempServer | Add-Member -Type NoteProperty -Name OutlookHTTP -Value 0
            $tempServer | Add-Member -Type NoteProperty -Name OutlookMAPIHTTP -Value 0
            if($AdminVersion -like "15*"){
                $OACounter = (Get-Counter -counter "\RPC/HTTP Proxy\RPC/HTTP Requests per Second" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
                if($Null -ne $OACounter){
                    $OACounter = $OACounter.toString()
                    if($OACounter.Contains("."))
                    {
                        $OACounter = $OACounter.Substring(0,$OACounter.IndexOf('.'))
                    }
                    [int]$OACounter = [convert]::ToInt32($OACounter, 10)
                    $totalOA += $OACounter            
                    $tempServer.OutlookHTTP = $OACounter
                }
        
                $MAPIHTTPCounter = (Get-Counter -counter "\MSExchange MapiHttp Emsmdb\Active User Count" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
                if($Null -ne $MAPIHTTPCounter){
                    $MAPIHTTPCounter = $MAPIHTTPCounter.toString()
                    if($MAPIHTTPCounter.Contains("."))
                    {
                        $MAPIHTTPCounter = $MAPIHTTPCounter.Substring(0,$MAPIHTTPCounter.IndexOf('.'))
                    }
                    [int]$MAPIHTTPCounter = [convert]::ToInt32($MAPIHTTPCounter, 10)
                    $totalMapiHTTP += $MAPIHTTPCounter
                    $tempServer.OutlookMAPIHTTP = $MAPIHTTPCounter
                }
            }
 
            $OWACounter = (Get-Counter -counter "\MSExchange OWA\Current Unique Users" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
            $tempServer | Add-Member -Type NoteProperty -Name OWA -Value $OWACounter
            $totalOWA += $OWACounter
 
            $EWSCounter = (Get-Counter -counter "\Web Service(_Total)\Current Connections" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
            $tempServer | Add-Member -Type NoteProperty -Name EWS -Value $EWSCounter
            $totalEWS += $EWSCounter
 
            $EASCounter = (Get-Counter -counter "\MSExchange ActiveSync\Requests/sec" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
            [string]$EASCounter = $EASCounter.toString()
                    if($EASCounter.Contains("."))
                    {
                        $EASCounter = $EASCounter.Substring(0,$EASCounter.IndexOf('.'))
                    }
                    [int]$EASCounter = [convert]::ToInt32($EASCounter, 10)
            $tempServer | Add-Member -Type NoteProperty -Name EAS -Value $EASCounter
            $totalEAS += $EASCounter
 
            $MaxCAPI = (Get-Counter -counter "\Netlogon(_total)\Semaphore Timeouts" -ComputerName $CAS.Name).CounterSamples[0].CookedValue
            $tempServer | Add-Member -Type NoteProperty -Name "Semaphore Timeouts" -Value $MaxCAPI
            $totalMaxCAPI += $MaxCAPI
            
            # Checking .NET version installed
            $regkey = Invoke-command -computername $CAS -scriptblock{Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | Select-Object Release}
            Switch($regkey.Release){
                378389 {$netver = ".NET Framework 4.5"}
                378675 {$netver =  ".NET Framework 4.5.1"}
                378758 {$netver =  ".NET Framework 4.5.1"}
                379893 {$netver =  ".NET Framework 4.5.2"}
                393295 {$netver =  ".NET Framework 4.6"}
                393297 {$netver =  ".NET Framework 4.6"}
                394254 {$netver =  ".NET Framework 4.6.1"}
                394271 {$netver =  ".NET Framework 4.6.1"}
                394802 {$netver =  ".NET Framework 4.6.2"}
                394806 {$netver =  ".NET Framework 4.6.2"}
                460798 {$netver =  ".NET Framework 4.7"}
                460805 {$netver =  ".NET Framework 4.7"}
                461308 {$netver =  ".NET Framework 4.7.1"}
                461310 {$netver =  ".NET Framework 4.7.1"}
                461808 {$netver =  ".NET Framework 4.7.2"}
                461814 {$netver =  ".NET Framework 4.7.2"}
                528040 {$netver =  ".NET Framework 4.8"}
                528049 {$netver =  ".NET Framework 4.8"}
                528209 {$netver =  ".NET Framework 4.8"}
            }
            $tempServer | Add-Member -Type NoteProperty -Name ".NET version" -Value $netver
        


            #checking if Visual C++ 2013 Redist is installed
            
            #Define the variable to hold the location of Currently Installed Programs
            $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

            #Create an instance of the Registry Object and open the HKLM base key
            $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$CAS.Name) 

            #Drill down into the Uninstall key using the OpenSubKey Method
            $regkey=$reg.OpenSubKey($UninstallKey) 

            #Retrieve an array of string that contain all the subkey names
            $subkeys=$regkey.GetSubKeyNames() 

            #Open each Subkey and use GetValue Method to return the required values for each
            $VC2013Redist = "Not Installed"
            $VC2012Redist = "Not Installed"
            foreach($key in $subkeys){

                $thisKey=$UninstallKey+"\\"+$key 
                $thisSubKey=$reg.OpenSubKey($thisKey) 
                
                if($thisSubKey.GetValue("DisplayName") -like "Microsoft Visual C++ 2013*"){
                    $VC2013Redist = "Installed"
                } 
                if($thisSubKey.GetValue("DisplayName") -like "Microsoft Visual C++ 2012*"){
                    $VC2012Redist = "Installed"
                } 
                
            }
            $tempServer | Add-Member -Type NoteProperty -Name "Visual C++ 2012 Redist" -Value $VC2012Redist
            $tempServer | Add-Member -Type NoteProperty -Name "Visual C++ 2013 Redist" -Value $VC2013Redist
            
            $ServersList += $tempServer
               
           
 
        #Clearing Variables
        $netver = $null
        $VC2013Redist = $null
        $OACounter
        $MAPIHTTPCounter
        $OWACounter
        $EWSCounter
        $EASCounter
        $MaxCAPI
        }
 
    }
        # Adding Totals to the last line of the output
        $tempServer = New-Object System.Object
        $tempServer | Add-Member -Type NoteProperty -Name ServerName -Value "Total"
        $tempServer | Add-Member -Type NoteProperty -Name OutlookRpc -Value $totalOutlook
        $tempServer | Add-Member -Type NoteProperty -Name OutlookHTTP -Value $totalOA
        $tempServer | Add-Member -Type NoteProperty -Name OutlookMAPIHTTP -Value $totalMapiHTTP
        $tempServer | Add-Member -Type NoteProperty -Name OWA -Value $totalOWA
        $tempServer | Add-Member -Type NoteProperty -Name EWS -Value $totalEWS
        $tempServer | Add-Member -Type NoteProperty -Name EAS -Value $totalEAS
        $tempServer | Add-Member -Type NoteProperty -Name "Semaphore Timeouts" -Value $totalMaxCAPI
 
        $ServersList += $tempServer
#endregion

#region Collect Server General Info
     }elseif($CASLoad -eq $false -and $CheckASACredentials -eq $false -and $Autodiscover -eq $false){
    # If no switch is used, we will only collect General Server information
    $Servers = Get-ExchangeServer
    # Check if "Server" parameter was in used
    if($Server -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.Name -like "*$Server*"} | Sort-Object Name
        }
        # Check if "Site" parameter was in used
        elseif($Site -ne "*"){
        $Servers = Get-ExchangeServer | Where-Object{$_.Site -like "*$Site*"} | Sort-Object Name
        Write-Host "Getting servers from Site: $Site" -ForegroundColor Green
    }
 
    $ServersList = @()
    foreach($CAS in $Servers){
        # Testing if server is reachable
        $conn = Test-Connection -ComputerName $CAS -Count 1 -ErrorAction SilentlyContinue
 
        # Getting info only if connection is alive
        if($null -ne $conn){
            Write-Host "Validating info against Server $CAS" -ForegroundColor Green
            $tempServer = New-Object System.Object
            $tempServer | Add-Member -Type NoteProperty -Name ServerName -Value $CAS
            . Get-ServerVersion -CAS $CAS
            $tempServer | Add-Member -Type NoteProperty -Name "ServerBuild          " -Value $exSetupVer
            $tempServer | Add-Member -Type NoteProperty -Name Site -Value $CAS.Site.Rdn.UnescapedName
		    $role = ""
            if($CAS.IsClientAccessServer -eq "True"){
                $role += "C"
            }
            if($CAS.IsHubTransportServer -eq "True"){
                $role += "H"
            }
            if($CAS.IsMailboxServer -eq "True"){
                $role += "M"
            }            
            if($CAS.IsUnifiedMessagingServer -eq "True"){
                $role += "U"
            }            
            if($CAS.IsEdgeServer -eq "True"){
                $role += "E"
            }                        
            $tempServer | Add-Member -Type NoteProperty -Name ServerRoles -Value $role

            # Getting .NET Version
            $regkey = Invoke-command -computername $CAS -scriptblock{Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | Select-Object Release}
            Switch($regkey.Release){
                378389 {$netver = ".NET Framework 4.5"}
                378675 {$netver =  ".NET Framework 4.5.1"}
                378758 {$netver =  ".NET Framework 4.5.1"}
                379893 {$netver =  ".NET Framework 4.5.2"}
                393295 {$netver =  ".NET Framework 4.6"}
                393297 {$netver =  ".NET Framework 4.6"}
                394254 {$netver =  ".NET Framework 4.6.1"}
                394271 {$netver =  ".NET Framework 4.6.1"}
                394802 {$netver =  ".NET Framework 4.6.2"}
                394806 {$netver =  ".NET Framework 4.6.2"}
                460798 {$netver =  ".NET Framework 4.7"}
                460805 {$netver =  ".NET Framework 4.7"}
                461308 {$netver =  ".NET Framework 4.7.1"}
                461310 {$netver =  ".NET Framework 4.7.1"}
                461808 {$netver =  ".NET Framework 4.7.2"}
                461814 {$netver =  ".NET Framework 4.7.2"}
                528040 {$netver =  ".NET Framework 4.8"}
                528049 {$netver =  ".NET Framework 4.8"}
                528209 {$netver =  ".NET Framework 4.8"}
            }
     
            $tempServer | Add-Member -Type NoteProperty -Name ".NET version" -Value $netver

            #checking if Visual C++ 2013 Redist is installed
            
            #Define the variable to hold the location of Currently Installed Programs
            $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

            #Create an instance of the Registry Object and open the HKLM base key
            $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$CAS.Name) 

            #Drill down into the Uninstall key using the OpenSubKey Method
            $regkey=$reg.OpenSubKey($UninstallKey) 

            #Retrieve an array of string that contain all the subkey names
            $subkeys=$regkey.GetSubKeyNames() 

            #Open each Subkey and use GetValue Method to return the required values for each
            $VC2013Redist = "Not Installed"
            $VC2012Redist = "Not Installed"
            foreach($key in $subkeys){

                $thisKey=$UninstallKey+"\\"+$key 
                $thisSubKey=$reg.OpenSubKey($thisKey) 
                
                if($thisSubKey.GetValue("DisplayName") -like "Microsoft Visual C++ 2013*"){
                    $VC2013Redist = "Installed"
                } 
                if($thisSubKey.GetValue("DisplayName") -like "Microsoft Visual C++ 2012*"){
                    $VC2012Redist = "Installed"
                } 
                
            }
            $tempServer | Add-Member -Type NoteProperty -Name "Visual C++ 2012 Redist" -Value $VC2012Redist
            $tempServer | Add-Member -Type NoteProperty -Name "Visual C++ 2013 Redist" -Value $VC2013Redist

            $ServersList += $tempServer
             
        }
 
    }
#endregion

}
# Out putting the results to the Console 
$ServersList | Format-Table -AutoSize
