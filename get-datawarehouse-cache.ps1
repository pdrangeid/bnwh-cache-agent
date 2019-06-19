<# 
.SYNOPSIS 
 PowerShell agent to collect data and submit to the datawarehouse via API
 
 
.DESCRIPTION 
 The webAPI will identify based on your tenantGUID which data sources, and time periods
 are being requested.  This agent will then query the data source(s), collect the 
 data and submit it via the WebAPI for submision to the data warehouse cache database.
 

.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ get-datawarehouse-cache.ps1                                                                 │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 6.17.2019 				               									  │ 
│   AUTHOR      : Paul Drangeid 			                   								  │ 
│   SITE        : https://github.com/pdrangeid/bnwh-cache-agent                               │ 
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#> 

param (
    [string]$subtenant,
    [boolean]$queryo365,
    [boolean]$noui
    )

$VMwareinitialized = $false
$ErrorActionPreference = 'SilentlyContinue'
Remove-Variable -name apikey | Out-Null
Remove-Variable -name tenantguid | Out-Null
$global:srccmdline= $($MyInvocation.MyCommand.Name)
$scriptappname = "Blue Net get-datawarehouse-cache"
$baseapiurl="https://api-cache.bluenetcloud.com"
$ScheduledJobName = "Blue Net Warehouse Data Refresh"
if (![string]::IsNullOrEmpty($subtenant)){
    $ScheduledJobName = "Blue Net Warehouse Data ($subtenant) Refresh"
}

Write-Host "`nLoading includes: $PSScriptRoot\bg-sharedfunctions.ps1"
Try{. "$PSScriptRoot\bg-sharedfunctions.ps1" | Out-Null}
Catch{
    Write-Warning "I wasn't able to load the sharedfunctions includes (which should live in the same directory as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
    Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\bg-sharedfunctions.ps1"
    BREAK
    }

    Prepare-EventLog
    Function Set-CacheSyncJob{

        Get-ScheduledTask -TaskName $ScheduledJobName -ErrorAction SilentlyContinue -OutVariable task
        if (!$task) {
        # task does not exist, otherwise $task contains the task object
        $answer=yesorno "Would you like to schedule this agent to run automatically?" "Schedule data synchronization"
        if ($answer -eq $true){
            $Username = (get-aduser ($Env:USERNAME)).userprincipalname
            $credentials = $Host.UI.PromptForCredential("Task username and password","Provide the password for this account that will run the scheduled task",$Username,$env:userdomain)
            $Password = $Credentials.GetNetworkCredential().Password 
            $Prog = $env:systemroot + "\system32\WindowsPowerShell\v1.0\powershell.exe"
            $thisuserupn = (get-aduser ($Env:USERNAME)).userprincipalname
            $Opt = "-nologo -noninteractive -noprofile -ExecutionPolicy BYPASS "+$PSScriptRoot+"\get-datawarehouse-cache.ps1 -noui $true"
            $Action = New-ScheduledTaskAction -Execute $Prog -Argument $Opt  -WorkingDirectory $PSScriptRoot
            $Trigger = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -At "01:00"
            $Trigger.Repetition = $(New-ScheduledTaskTrigger -Once -At "02:00" -RepetitionDuration "22:00" -RepetitionInterval "00:10").Repetition
            $Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 10 -StartWhenAvailable
            $Settings.ExecutionTimeLimit = "PT30M"
            Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -TaskName $ScheduledJobName -Description "Periodically sends updated data to the reporting datawarehouse via WebAPI" -User $Username -Password $Password -RunLevel Highest
        }# Yes - operater wants us to schedule this task
            }# End if (task doesn't already exist)

        }#End Function

    Function init-vmware(){
    # Ensure we are able to load the get-vmware-data.ps1 include.
        $ErrorActionPreference = 'stop'
        Write-host "loading the vmware include file...$PSScriptRoot\get-vmware-data.ps1"
        #Try{. "$PSScriptRoot\get-vmware-data.ps1" | Out-Null}
        Try{. "$PSScriptRoot\get-vmware-data.ps1"}
        Catch{
            Write-Warning "I wasn't able to load the get-vmware-data.ps1 include script (which should live in the same directory as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
            Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\get-vmware-data.ps1"
            BREAK
            }# End Catch
            return $true
    }#End init-vmware function
    Function init-adsi(){
    # Verify we can load the Active Directory module.  If not prompt to download and install
    $ErrorActionPreference = 'Stop'
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
                return $false
            }

            #If there are old domain controllers (or not running AD Web Services) you can skip them by adding their hostname to the 'skipdc' reg_sz value
            $Path = "HKCU:\Software\BNCacheAgent"
            $dcskiplist=Ver-RegistryValue -RegPath $Path -Name "skipdc" -DefValue "Skipthisserver"
            $dcskiplist = if ($dcskiplist -eq $null) { "Skipthisserver" } else { $dcskiplist}
            if (! $dcskiplist -eq 'Skipthisserver') {write-host "per registry config Skipping $dcskiplist"}
            Do {
                $serverlist=netdom query dc| ForEach-Object{
                    if (![string]::IsNullOrEmpty($_) -and $_ -notmatch "command completed" -and $_ -notmatch "List of domain" -and $_.toLower() -notmatch $dcskiplist ) {
                        if (![string]::IsNullOrEmpty($global:targetserver)) {
                            return}
                    Write-Host "`nAttempt to query ActiveDirectory via $_"
                    $tenantname = get-addomain -server $_ | select -ExpandProperty "name"
                    Write-Host "`nIdentified the tenantdomain as: '$tenantname'"
                    if (![string]::IsNullOrEmpty($tenantname)) {
                        Write-Host "Setting target Domain Controller to $_"
                        $global:targetserver=$($_)
                    }# endif tenantname not null
                    }# this DC is a non-skip DC
                    }
                    write-host "now a break?"
                    
            }
            until (![string]::IsNullOrEmpty($global:targetserver))
            Write-Host "The target Domain Controller is $global:targetserver"
        
        $tenantdomain = get-addomain -server $targetserver| select -ExpandProperty "DNSRoot"
        $shortdomain = $tenantdomain.replace('.','_')
        return $true
        }# End init-adsi function
    
    Function submit-cachedata($Cachedata,[string]$DSName){
    # Takes the resulting cachedata and submits it to the webAPI 
        Write-Host "Submitting Data for $DSName"
        #write-host "******************************* the cache data is: `n"$Cachedata
        $ErrorActionPreference = 'Stop'
        Try{
        $apibase="https://api-cache.bluenetcloud.com/api/v1/submit-data/"
        $apiurlparms="?TenantGUID="+$tenantguid+"&DataSourceName="+$DSName+"&NewTimeStamp="+$querytimestamp
        $apiurl=$apibase+$apiurlparms.replace('+','%2b')
        $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint($apiurl)
        $thecontent = @{"data" = $Cachedata}
        #$pjson = $($params | ConvertTo-Json -Depth 5 -Compress
        #$thecontent = $(@{"data" = $Cachedata} | ConvertTo-Json -Compress)
        $thecontent = (@{"data" = $Cachedata} | ConvertTo-Json)
        $ErrorActionPreference= 'silentlycontinue'
        $pjmb=[math]::Round(([System.Text.Encoding]::UTF8.GetByteCount($Cachedata))*0.00000095367432,2) 
        write-host "Submitting $ic updates for $DSName ($([math]::Round($pjmb,2))MB)"
        if ($DSName -like '*vmware*'){
            $thecontent = $Cachedata
            Invoke-RestMethod $apiurl -Method 'Post' -Headers @{"x-api-key"=$APIKey;"content-type" = "binary"} -Body $thecontent -ErrorVariable RestError -ErrorAction SilentlyContinue -TimeoutSec 900
            }
            else {
        Invoke-RestMethod $apiurl -Method 'Post' -Headers @{"x-api-key"=$APIKey;Accept="application/json";"content-type" = "binary"} -Body $thecontent -ErrorVariable RestError -ErrorAction SilentlyContinue -TimeoutSec 900
        write-host "******************************* the body data is: `n"$thecontent
            }
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
    function get-o365admin([boolean]$allowpwchange){
        Add-Type -AssemblyName Microsoft.VisualBasic
        $Path = "HKCU:\Software\BNCacheAgent\$subtenant\o365"
        $Path=$path.replace('\\','\')
        AddRegPath $Path
        write-host "checking $Path"
        $result = Get-Set-Credential "Office365" $Path "o365AdminUser" "o365AdminPW" $false "administrator@company.com"
        $credUser = Ver-RegistryValue -RegPath $Path -Name "o365AdminUser"
        $credPwd = Ver-RegistryValue -RegPath $Path -Name "o365AdminPW"
        $securePwd = ConvertTo-SecureString $credPwd
        $global:o365cred = New-Object System.Management.Automation.PsCredential($credUser, $securePwd)
        Try{
        Connect-MsolService -Credential $o365cred
        }
        Catch {
            write-host "failed to verify credentials and/or connect to the MsolService"
            Write-Host "returning false"
            return $false
        }
        Write-Host "returning true"
        return $true
    }#End Function (get-o365admin)

    Function get-o365-assets([string]$objclass){
        Write-host "getting o365 assets"
            $ErrorActionPreference = 'Stop'
        # -----------------------------------------------------
        # Set parameters for vCenter and start RVTools export
        # -----------------------------------------------------
        $Path = "HKCU:\Software\BNCacheAgent\$subtenant\o365"
        $Path = $Path.replace('\\','\')
        write-host "Delegated Admin is $O365Delegated"
            Write-Host "Using supplied authentication credentials"
            Write-Host "Using supplied authentication username:"$o365cred.username
            Connect-MsolService -Credential $o365cred
            write-host "The objclass is $objclass"

            if ($objclass -like '*user'){
            $o365results=(Get-MsolUser | Select-Object * )
            }

            elseif ($objclass -like '*device'){
                $o365results=(Get-MsolDevice -All | Select-Object *)
            }
    
            elseif ($objclass -like '*contact'){
                $o365results=(Get-MsolContact -All | Select-Object *)
            }
    
            elseif ($objclass -like '*accountsku'){
                $o365results=(Get-MsolAccountSku | Select-Object *)
            }

            elseif ($objclass -like '*group'){
                $o365results=(Get-MsolGroup | Select-Object *)
            }

            elseif ($objclass -like '*licensetype'){
                $o365results=(Get-MsolUser -All | Select DisplayName,userPrincipalname,isLicensed,BlockCredential,ValidationStatus,@{n="Licenses Type";e={$_.Licenses.AccountSKUid}})
            }

            elseif ($objclass -like '*mailbox'){
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365cred -Authentication  Basic -AllowRedirection
                Import-PSSession $Session -DisableNameChecking
                #$o365results=(Get-MsolUser -All | Where-Object {$_.IsLicensed -eq $true -and $_.BlockCredential -eq $false} | Select-Object UserPrincipalName | ForEach-Object {Get-Mailbox -Identity $_.UserPrincipalName | Where-Object {$_.WhenChangedUTC -ge $tenantlastupdate} | Select-Object *})
                $o365results=(Get-Mailbox | Where-Object {$_.WhenChangedUTC -ge $tenantlastupdate} | Select-Object *)
                Remove-PSSession $Session
            }

            else {
                write-host "We got something we didn't quite expect..."
                write-host "request for $objclass"
                return
            }

            $ic = [int]($o365results | measure).count
            write-host "We got $ic results for $objclass"
            Write-host "Assuming all went well, Now do some processing and uploading..."
            return  $($o365results)
        }

Function get-mailprotector(){

}

Function get-filteredadobject([string]$ADObjclass,[string]$requpdate){
    $ErrorActionPreference = 'stop'
    $DefDate = 	[datetime]"4/25/1980 10:05:50 PM"
    $global:querytimestamp=[DateTime]::UtcNow | get-date -Format "yyyy-MM-ddTHH:mm:ss"
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
    Try{
    if (![string]::IsNullOrEmpty($mysearchbase)){
        write-host "let's split up the searchbase"
        $arrsb=@($mysearchbase -split '\r?\n')# If the regvalue was multi-line we need to split it into multiple searchbase entries
        $adresults=($arrsb | ForEach {Get-ADObject -resultpagesize 50 -server $targetserver -Searchbase $_ -Filter $myfilter -Properties * -ErrorAction SilentlyContinue})
        }#We have a custom searchbase
    else {
        write-Host "AD Query: Get-ADObject -resultpagesize 50 -server $targetserver -Filter $myfilter -Properties *"
        $adresults = Get-ADObject -resultpagesize 50 -server $targetserver -Filter $myfilter -Properties *
        }# No custom searchbase
    }#End Try

    Catch{
        Write-Host "Sorry - we failed miserably"
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host "message: $ErrorMessage  item:$FailedItem"
        if ($ErrorMessage -like "*object not found*"){
            Write-Warning "Possibly a permissions issue with the user account querying Active Directory?"
        }
        exit
    } # End Catch

    $ic = [int]($adresults | measure).count
    Write-Host "Found $ic $ADObjclass updates."
    if ($ic -eq 0) {
    $adoutput = "We have 0 results"
    }
    if ($ic -ge 1) {
    $allProperties =  $adresults | %{ $_.psobject.properties | select Name } | select -expand Name -Unique | sort  
    $adoutput = $adresults | select $allProperties}#We had at least 1 result in $ic
    Write-Host "We got $ic $ADObjclass updates to submit to the API."
    submit-cachedata $adoutput $($_.SourceName)
    write-host "did we submit the data?"
    return
}# End Function get-filteredadobject

$Path = "HKCU:\Software\BNCacheAgent\$subtenant\"
    $Path=$Path.replace('\\','\')

    Try{
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
        $APIKey = GetKey $($Path) $("APIKey") $("Enter global APIKey in the password field:")
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host "Failed to retrieve APIKey from registry The error message was '$ErrorMessage'  It is likely that you are not running this script as the original user who saved the APIKey value."
        Break
        exit
    }

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
    Write-Host "Req: "($R2.DataRequests | measure).count
    exit
    }
  
    $o365req=($r2.DataRequests | where-object -Property SourceName -like 'O365*')
    if (![string]::IsNullOrEmpty($o365req) -and $queryo365 -eq $true){
        Try{
            write-host "Let's init o365"
            $o365initialized=(get-o365admin $false)
            write-host "got o365admin  and the results are $o365initialized"
        }
        Catch{
            write-host "Sorry - O365 init epic failure!"
        }
        
    }# end initializing O365

    $vmwarereq=($r2.DataRequests | where-object -Property SourceName -like '*vmware*')
    if (![string]::IsNullOrEmpty($vmwarereq)){
    
    Write-host "loading the vmware include file...$PSScriptRoot\get-vmware-data.ps1"
    #Try{. "$PSScriptRoot\get-vmware-data.ps1" | Out-Null}
    Try{. "$PSScriptRoot\get-vmware-data.ps1"}
    Catch{
         Write-Warning "I wasn't able to load the get-vmware-data.ps1 include script (which should live in the same directory as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
         Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\get-vmware-data.ps1"
         BREAK
    }
    
    $VMwareinitialized=(get-vcentersettings $false)
    write-host "got vmsettings and the results are $VMwareinitialized"
    }

    $adsireq=($r2.DataRequests | where-object -Property SourceName -like 'ADSI*')
    if (![string]::IsNullOrEmpty($adsireq)){
        $ErrorActionPreference = 'Stop'
        Try{
            write-host "Let's init ActiveDirectory"
            $adinitialized=(init-adsi)    
        }
        Catch{
            write-host "Sorry - ActiveDirectory initialization was an epic failure!"
        }
        write-host "got init-adsi and the results are $adinitialized"
    }# end initializing AD Module
    
$dr = 0
Write-Host "Processing "$(($R2.DataRequests | measure).count) "data object requests."
$R2.DataRequests | ForEach-Object{
$dr++
Write-Host "Processing $dr of"$(($R2.DataRequests | measure).count) "data object requests."
$global:querytimestamp=[DateTime]::UtcNow | get-date -Format "yyyy-MM-ddTHH:mm:ss"
$ModDate=$_.NextUpdate
$MaxAge=$_.MaxAgeMinutes
$HasModified=$_.HasModifiedDate
$Delegated=$_.O365DelegatedAdmin
$SourceReqUpdate = $false

if ($querytimestamp -ge $ModDate) {
   $SourceReqUpdate=$true
   Write-Host $_.SourceName "Next Update requested at/after [$ModDate] with a MaxAge of $MaxAge and will be updated."
}
if (!$SourceReqUpdate){
    Write-Host $_.SourceName "Next Update requested at/after [$ModDate] with a MaxAge of $MaxAge and is not in need of a query"
    return
}
    if ($_.SourceName -like "*ADSI*"){
        if ($adinitialized -eq $false){
            Write-Warning "API has requested Active Directory data, but I could not initialize the ActiveDirectory Module."
            exit
            }
    $Source=$_.SourceName.replace('ADSI-','')
    #Write-Host "Request for Active Directory $Source data from $ModDate or later."
    $ErrorActionPreference = 'Stop'
    $intresult=(get-filteredadobject $($Source) $($ModDate))
    write-host "We got about "$intresult "items returned"
    }# end if (ADSI source request)
elseif ($_.SourceName -like "*vmware*"){
    $Source=$_.SourceName.replace('VMware ','')
    if ($VMwareinitialized -eq $false){
        Write-Warning "API has requested VMware data, but I could not initialize the VMware data requester."
        exit
        }
    if ($VMwareinitialized -eq $true){
        Write-host "We're gonna get VM assets ("$_.SourceName") for $Source"
        $global:querytimestamp=[DateTime]::UtcNow | get-date -Format "yyyy-MM-ddTHH:mm:ss"
        $vmresult=get-vmware-assets $Source
        write-host "The resulting VM data request is..."
        Write-host "vmr: $vmresult"
        if ($vmresult -ne $false){
            #Now take the resulting export file and submit to the cache ingester:
            Get-ChildItem $vmresult -Filter *.csv | Foreach-Object { 
                $Objclass = $($_.Name).replace('RVTools_tab','').replace('.csv','')
                #Write-Host "Let's send VM Data ($Source) from $Objclass to the API Cache ingester!"
                $csvfilename = "$vmresult\"+$_.Name
                #$content = (Import-Csv -Path $csvfilename)
                $content = [IO.File]::ReadAllText($csvfilename);
                $ic=(Import-Csv $csvfilename | Measure-Object).count
                $srcname="Vmware "+$Source
                submit-cachedata $content $srcname
                write-host "and here's the data we will submit `n $content"
                Remove-Item -path $csvfilename
            }# end Foreach-Object
            Remove-Item -path $vmresult -Recurse
        } # End if we received a valid VM data export file!
    } # end if VMwareinistialized
} # End elseif $_.SourceName -like "*vmware*"
elseif ($_.SourceName -like "o365*"){

if ($queryo365 -ne $true){
    Write-Host "Office365 processing was not enabled.  to enable, add -queryo365 $true to the commandline when executing the script."
    return
}

if ($o365initialized -eq $false){
    Write-Warning "API has requested O365 data, but I could not initialize the MsolService."
    exit
    }
    $global:querytimestamp=[DateTime]::UtcNow | get-date -Format "yyyy-MM-ddTHH:mm:ss"
    $o365result=get-o365-assets $_.Sourcename
    submit-cachedata $o365result $_.SourceName

}# $_.SourceName -like "o365*"


else {
    write-host "Some other data request... "$_.SourceName" ... and I have no idea what to do with it!"
    return
}
}# Next $R2.DataRequests object
Write-Host "All Done processing "$(($R2.DataRequests | measure).count) " requests."
Get-PSSession | Remove-PSSession

if ($noui -ne $true){
    # Check to see if the job is scheduled
    Set-CacheSyncJob
}

Remove-Variable -name apikey | Out-Null
Remove-Variable -name tenantguid | Out-Null
Remove-Variable -name params | Out-Null





