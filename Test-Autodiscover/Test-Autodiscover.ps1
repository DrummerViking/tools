<#
    .SYNOPSIS
    Function to test Autodiscover V2.
    
    .DESCRIPTION
    Function to test Autodiscover V2 against an on-premises server or Office 365. You can also select different protocols available.
    
    .PARAMETER EmailAddress
    Email address of the user account you want to test.
    
    .PARAMETER Server
    This is an optional parameter. In case you want to specifically test AutodiscoverV2 against an on-premises FQDN or Office 365. The default value is "Outlook.office365.com".
    
    .PARAMETER Protocol
    Select one of the following mandatory values: "AutodiscoverV2","ActiveSync","Ews","Rest","Substrate","SubstrateNotificationService","SubstrateSearchService","OutlookMeetingScheduler".
    
    .PARAMETER ShowQueriedUrl
    This is an optional parameter. It will show the QueriedUrl in case you want to copy and paste into a browser.
    
    .EXAMPLE
    PS C:\> Test-Autodiscover -EmailAddress onpremUser@contoso.com -Protocol AutodiscoverV2 -ShowQueriedUrl
    In this example it will show the autodiscover URL for the onpremises user, queried against outlook.office365.com

    .EXAMPLE
    PS C:\> Test-Autodiscover -EmailAddress cloudUser@contoso.com -Protocol EWS -Server mail.contoso.com
    In this example it will show the EWS URL for the cloud user, queried against an on-premises endpoint 'mail.contoso.com'.

    #>
Param (
    [Parameter( Mandatory = $true, Position = 0)]
    [String]$EmailAddress,
 
    [Parameter( Mandatory = $false, Position = 1)]
    [String]$Server = "outlook.office365.com",
 
    [Parameter( Mandatory = $true, Position = 2)]
    [ValidateSet("AutodiscoverV2", "ActiveSync", "Ews", "Rest", "Substrate", "SubstrateNotificationService", "SubstrateSearchService", "OutlookMeetingScheduler")]
    [String]$Protocol,

    [Parameter( Mandatory = $false, Position = 3)]
    [Switch]$ShowQueriedUrl
 
)
if ($Protocol -eq "AutodiscoverV2") { $protocolUsed = "AutodiscoverV1" }
else { $protocolUsed = $Protocol }
    
try {
    $URL = "https://$server/autodiscover/autodiscover.json?Email=$EmailAddress&Protocol=$protocolUsed&RedirectCount=5"
    Write-Verbose "URL=$($Url)"

    $response = Invoke-RestMethod -Uri $Url -UserAgent Teams 

    if ( $ShowQueriedUrl ) {
        [PSCustomObject]@{
            User          = $EmailAddress
            QueriedServer = $Server
            Protocol      = $protocolUsed
            ReturnedUrl   = $response.URL
            QueriedURL    = $URL
        } | Format-List
    }
    else {
        [PSCustomObject]@{
            User          = $EmailAddress
            QueriedServer = $Server
            Protocol      = $protocolUsed
            ReturnedUrl   = $response.URL
        } | Format-List
    }
}
catch {
    #create object
    $returnValue = New-Object -TypeName PSObject
    #get all properties from last error
    $ErrorProperties = $Error[0] | Get-Member -MemberType Property
    #add existing properties to object
    foreach ($Property in $ErrorProperties) {
        if ($Property.Name -eq 'InvocationInfo') {
            $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
        }
        else {
            $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
        }
    }
    #return object
    $returnValue
    break
}