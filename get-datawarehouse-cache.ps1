<# 
.SYNOPSIS 
 PowerShell agent to collect data and submit to the datawarehouse via API
 
 
.DESCRIPTION 
 The webAPI will identify based on your tenantGUID which data sources, and time periods
 are being requested.  This agent will then query the local data sources, collect the 
 data and submit it via the WebAPI for submision to the data warehouse cache database.
 

.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ get-datawarehouse-cache.ps1                                                                 │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 5.28.2019 				               									  │ 
│   AUTHOR      : Paul Drangeid 			                   								  │ 
│   SITE        : https://github.com/pdrangeid/bnwh-cache-agent                               │ 
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#> 

param (
    [string]$subtenant
    )

$VMwareinitialized = $false
$ErrorActionPreference = 'SilentlyContinue'
Remove-Variable -name apikey | Out-Null
Remove-Variable -name tenantguid | Out-Null
$global:srccmdline= $($MyInvocation.MyCommand.Name)
$scriptappname = "Blue Net get-datawarehouse-cache"
$=0
#Write-Host "I live in $PSScriptRoot"

Write-Host "`nLoading includes: $PSScriptRoot\bg-sharedfunctions.ps1"
Try{. "$PSScriptRoot\bg-sharedfunctions.ps1" | Out-Null}
Catch{
    Write-Warning "I wasn't able to load the sharedfunctions includes (which should live in the same direcroty as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
    Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\bg-sharedfunctions.ps1"
    BREAK
    }

    Prepare-EventLog

    Function set-syncschedule(){
    }

    Function init-vmware(){
        Try{. "$PSScriptRoot\get-vmware-data.ps1" | Out-Null}
        Catch{
            Write-Warning "I wasn't able to load the get-vmware-data.ps1 include script (which should live in the same directory as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
            Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\get-vmware-data.ps1"
            BREAK
            }# End Catch
    }#End init-vmware function
    
    $baseapiurl="https://api-cache.bluenetcloud.com"

    $m = Get-Module -List activedirectory
    if(!$m) {
    $message1="Unable to find the ActiveDirectory PowerShell module.  This is required for operation.  For help please visit: " + "https://blogs.technet.microsoft.com/ashleymcglone/2016/02/26/install-the-active-directory-powershell-module-on-windows-10/  or https://www.google.com/search?q=how+to+install+the+Active+Directory+powershell+module"

    $answer=yesorno "Would you like the ActiveDirectory PowerShell module installed on this workstation?" "Missing AD Powershell Module"
    write-host $answer
    if ($answer -eq $true){
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $osInfo.ProductType
        if ($osInfo.ProductType -ne 1){
        Install-WindowsFeature RSAT-AD-PowerShell
        Write-Host "If the installation was successful, please try running the script again.  You SHOULD NOT require a reboot."
        exit
    } # Windows Server detected - use the Install-WindowsFeature method to install the AD tools
        elseif ( $((Get-WMIObject win32_operatingsystem).name) -like 'Microsoft Windows 10*' ) {
        #Write-Host "Download https://gallery.technet.microsoft.com/Install-the-Active-fd32e541/file/149000/1/Install-ADModule.p-s-1.txt"
    $client = new-object System.Net.WebClient
    $dwnloaddst = $env:temp+"\install-admodule.ps1"
    $client.DownloadFile("https://gallery.technet.microsoft.com/Install-the-Active-fd32e541/file/149000/1/Install-ADModule.p-s-1.txt",$dwnloaddst)
    if (Test-Path $dwnloaddst) {
    Write-Host "Installing ADModule...`n"
    Invoke-Expression "& `"$dwnloaddst`" "
    Write-Host "If the installation was successful, please try running the script again.  You SHOULD NOT require a reboot."
    exit
    } else {write-host "Download failed... You must install the ActiveDirectory PowerShell module for this agent to run properly.";
    } # Windows 10 detected
        } # User answered "yes, please install"
    } # We couldn't find the AD module installed
    
    Write-Warning $message1
    Sendto-eventlog -message $message1 -entrytype "Warning"
    BREAK
    }
    
    TRY{
        import-module activedirectory
    }
    CATCH{
        $message1="Unable to load the ActiveDirectory PowerShell module.  This is required for operation.  For help please visit: " + "https://blogs.technet.microsoft.com/ashleymcglone/2016/02/26/install-the-active-directory-powershell-module-on-windows-10/  or https://www.google.com/search?q=how+to+install+the+Active+Directory+powershell+module"
        Write-Warning $message1
        Sendto-eventlog -message $message1 -entrytype "Warning"
    BREAK
    }

    #If there are old domain controllers (or not running AD Web Services) you can skip them by adding their hostname to the 'skipdc' reg_sz value
    $Path = "HKCU:\Software\BNCacheAgent\"
    $dcskiplist=$(Ver-RegistryValue -RegPath $Path -Name "skipdc").toLower()
    $dcskiplist = if ($dcskiplist -eq $null) { "Skipthisserver" } else { $dcskiplist
        write-host "per registry config Skipping $dcskiplist"}
    Do {
        $serverlist=netdom query dc| ForEach-Object{
            if (![string]::IsNullOrEmpty($_) -and $_ -notmatch "command completed" -and $_ -notmatch "List of domain" -and $_.toLower() -notmatch $dcskiplist ) {
               Write-Host "`nAttempt to query ActiveDirectory via $_"
               $tenantname = get-addomain -server $_ | select -ExpandProperty "name"
               Write-Host "`nIdentified the tenantdomain as: '$tenantname'"
               if (![string]::IsNullOrEmpty($tenantname)) {
                   $targetserver=$($_)
                   Write-Host "Setting target Domain Controller to $_"
                BREAK
                }
            }
            }
            BREAK
    }
    until (![string]::IsNullOrEmpty($targetserver))
    
    Write-Host "The target Domain Controller is $targetserver"
    
    $tenantdomain = get-addomain -server $targetserver| select -ExpandProperty "DNSRoot"
    $shortdomain = $tenantdomain.replace('.','_')

    #write-host "Tenantdomain is $tenantdomain"
    #write-host "shortdomain is $shortdomain"

    Try
    {
        $Path = "HKCU:\Software\BNCacheAgent\"
        if (![string]::IsNullOrEmpty($subtenant)){$Path=$($Path+$subtenant+"\")}
        Write-host "the tenant guid storage path is $Path"
    $tenantguid = GetKey $($Path+$tenantdomain) $("TenantGUID") $("Enter Unique GUID for $tenantdomain in the password field:")
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host "Failed to retrieve tenant GUID from registry The error message was '$ErrorMessage'  It is likely that you are not running this script as the original user who saved the secure tenant GUID."
        Break
        exit
    }

    Try
    {
        $Path = "HKCU:\Software\BNCacheAgent\"
    $APIKey = GetKey $($Path+$tenantdomain) $("APIKey") $("Enter global APIKey in the password field:")
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host "Failed to retrieve tenant GUID from registry The error message was '$ErrorMessage'  It is likely that you are not running this script as the original user who saved the secure tenant GUID."
        Break
        exit
    }

    Function submit-cachedata($Cachedata,[string]$DSName){
        $ErrorActionPreference = 'stop'
        Try{
        #$apiurl="https://api-cache.bluenetcloud.com/api/v1/submit-data"
        $apiurl="https://api-cache.bluenetcloud.com/api/v1/submit-data/?TenantGUID="+$tenantguid+"&DataSourceName="+$DSName
        $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint($apiurl)
        #$params = @{"TenantGUID"=$tenantguid; "DataSourceName" = $DSName; "moredata" = @($Cachedata)}
        $params = @{"data" = $Cachedata}
        $pjson = $($params | ConvertTo-Json -Depth 5 -Compress)
        $pjmb=[math]::Round(([System.Text.Encoding]::UTF8.GetByteCount($pjson))*0.00000095367432,2)
        write-host "Submitting $m updates for $DSName ($([math]::Round($pjmb,2))MB)"
        Invoke-RestMethod $apiurl -Method 'Post' -Headers @{"x-api-key"=$APIKey;Accept="application/json";"content-type" = "binary"} -Body $pjson -ErrorVariable RestError -ErrorAction SilentlyContinue -TimeoutSec 900
        }
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            $httpresponse = $_.Exception.Response
            $HttpStatusCode = $RestError.ErrorRecord.Exception.Response.StatusCode.value__
            $HttpStatusDescription = $RestError.ErrorRecord.Exception.Response.StatusDescription
            write-host "Error Message $ErrorMessage `nFailed Item:$FailedItem `nhttp Response:$httpresponse`n"
            $result = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($result)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd();
            Write-Host "`nFailed to submit $m to $apiURL $ErrorMessage $FailedItem" -ForegroundColor Yellow
            Write-Host "HTTP Response Status Code: "$HttpStatusCode
            Write-Host "HTTP Response Status Description: "$HttpStatusDescription
            Write-Host "TenantName: "$TenantName
            Write-Host "Result: "$responseBody
            EXIT
        }
           
    }

Function get-filteredadobject([string]$ADObjclass,[string]$requpdate){
    $ErrorActionPreference = 'stop'
    $DefDate = 	[datetime]"4/25/1980 10:05:50 PM"
    $dtenow = (Get-Date).tostring()
    if ($requpdate -eq [DBNull]::Value -or [string]::IsNullOrEmpty($requpdate)) {
        $requpdate = [datetime]$DefDate
    }
        #Pull all the registry settings into a hashtable
        
        $Path = "HKCU:\Software\BNCacheAgent\"
        if (![string]::IsNullOrEmpty($subtenant)){$Path=$($Path+$subtenant)}
        $adsiconfigitems=(Get-Item $Path |
        Select-Object -ExpandProperty property |
        ForEach-Object {
        New-Object psobject -Property @{"property"=$_;
        "Value" = (Get-ItemProperty -Path $Path -Name $_).$_}})
    
    #To access a value from $adsiconfigitems
    # $myvalue=($adsiconfigitems | where-object -Property property -like 'Searchbase-objectclass').value
    $defsearchbase=($adsiconfigitems | where-object -Property property -like 'searchbase').value # Use this SearchBase value unless a more specific one is provided
    $matchstring=$("searchbase-"+$ADObjclass)#object specific searchbase will be 'searchbase-[objectclass]'.  You can provide multiple searchbases by using a REG_MULTI_SZ value
    $specificsearchbase=($adsiconfigitems | where-object -Property property -like $matchstring).value
    $mysearchbase=""
    if (![string]::IsNullOrEmpty($defsearchbase)) {$mysearchbase=$defsearchbase}#use default searchbase if it is defined
    if (![string]::IsNullOrEmpty($specificsearchbase)) {$mysearchbase=$specificsearchbase}# use an objectclass specific searchbase if it is defined
    $tenantlastupdate = [datetime]$requpdate
    write-output "`nAPI Requesting $ADObjclass data newer than [$tenantlastupdate]"
    $filtervalue = "modified -gt '" + $tenantlastupdate + "'"
    $ldapfilter = "(&(objectClass='$ADObjclass'))"
    $myfilter="(objectClass -eq '$ADObjclass') -and (modified -gt '$tenantlastupdate')"
    if (![string]::IsNullOrEmpty($mysearchbase)){
        write-host "let's split up the searchbase"
    $arrsb=@($mysearchbase -split '\r?\n')# If the regvalue was multi-line we need to split it into multiple searchbase entries
    Try{
    $adresults=($arrsb | ForEach {Get-ADObject -resultpagesize 50 -server $targetserver -Searchbase $_ -Filter $myfilter -Properties * -ErrorAction SilentlyContinue})
    }#End Try
    
    Catch{Write-Host "Sorry - we failed miserably"
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Host "message: $ErrorMessage  item:$FailedItem"
    exit} # End Catch
    }#We have a custom searchbase

    else {
    write-Host "AD Query: Get-ADObject -resultpagesize 50 -server $targetserver -Filter $myfilter -Properties *"
    $adresults = Get-ADObject -resultpagesize 50 -server $targetserver -Filter $myfilter -Properties *
    }# No custom searchbase

    $m = [int]($adresults | measure).count
    Write-Host "Found $m $ADObjclass updates."
    if ($m -ge 1) {
    $allProperties =  $adresults | %{ $_.psobject.properties | select Name } | select -expand Name -Unique | sort  
    $adoutput = $adresults | select $allProperties
    Write-Host "We got $m $ADObjclass updates to submit to the API."
    submit-cachedata $adoutput $($_.SourceName)
    write-host "did we submit the data?"
    }#We had at least 1 result in $m
    
}# End Function get-filteredadobject

#$APIKey
#$tenantguid
Try{
# Attempt to query the API to find out what data they would like us to retrieve
$Howsoonisnow=[DateTime]::UtcNow | get-date -Format "yyyy-MM-ddTHH:mm:ss"
$apiurl="https://api-cache.bluenetcloud.com/api/v1/get-data-requests"
$ServicePoint = [System.Net.ServicePointManager]::FindServicePoint($apiurl)
$params = @{"TenantGUID"=$tenantguid; "ClientAgentUTCDateTime" = $Howsoonisnow}
$Response = Invoke-RestMethod -uri $apiurl -Body $params -Method GET -Headers @{"x-api-key"=$APIKey;Accept="application/json"} -ErrorVariable RestError -ErrorAction SilentlyContinue
$Response | fl
}

Catch{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    $httpresponse = $_.Exception.Response
    $HttpStatusCode = $RestError.ErrorRecord.Exception.Response.StatusCode.value__
    $HttpStatusDescription = $RestError.ErrorRecord.Exception.Response.StatusDescription
    if ($ErrorMessage -eq 'Unable to connect to the remote server'){
        Write-Host "`n"
        Write-Warning "Unable to connect to the remote server $baseapiurl"
        Write-Host "Please check DNS, firewall, and Internet connectivity to verify."
        exit
    }
    write-host "Error Message $ErrorMessage `nFailed Item:$FailedItem `nhttp Response:$httpresponse`n"
    $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
    Write-Host "`nFailed to submit $m to $apiURL $ErrorMessage $FailedItem" -ForegroundColor Yellow
    Write-Host "HTTP Response Status Code: "$HttpStatusCode
    Write-Host "HTTP Response Status Description: "$HttpStatusDescription
    Write-Host "TenantName: "$TenantName
    Write-Host "Result: "$responseBody
    EXIT
}

$R2 = $Response | Convertfrom-Json
($R2.DataRequests | measure).count
if (($R2.DataRequests | measure).count -eq 0){
    Write-Host "Client identified successfully - no data requests at this time.  If this is your first run, please be sure the client's reporting setup has been completed."
    exit
    }
    
$R2.DataRequests | ForEach-Object{
if ($_.SourceName -like "*ADSI*"){
$Source=$_.SourceName.replace('ADSI-','')
$ModDate=$_.LastUpdate
Write-Host "Request for Active Directory $Source data from $ModDate or later."
get-filteredadobject $($Source) $($ModDate)
}# end if (ADSI source request)
elseif ($_.SourceName -like "*vmware*"){
    $Source=$_.SourceName.replace('VMware ','')
    if ($VMwareinitialized -eq $false){
        init-vmware
        $VMwareinitialized=get-vcentersettings $false
        write-host "got vmsettings and the results are $result"
        }
    if ($VMwareinitialized -eq $true){
        $vmresult=get-vmware-assets $Source
        }

    }
}
else {
    write-host "Some other data request... $_.SourceName"
}
}# Next $R2.DataRequests object

Remove-Variable -name apikey | Out-Null
Remove-Variable -name tenantguid | Out-Null
Remove-Variable -name params | Out-Null

Function get-vmwaredataobject ([string]$Objclass){
    
    if ($vmresult -ne $false){
        Get-ChildItem $vmresult -Filter *.csv | Foreach-Object { 
            $Objclass = $($_.Name).replace('RVTools_tab','').replace('.csv','')
            Write-Host "Let's send VM Data $Objclass to the API Cache ingester!"
            $content = Import-Csv -Path $_.Name
            submit-cachedata $Cachedata $Objclass
            Remove-Item -path $_.Name
        }# end Foreach-Object
        Remove-Item -path $vmresult -Recurse
    } # End if we got valid VM data files!
    exit
    
}
