<# 
.SYNOPSIS 
 PowerShell agent to collect MWP device data and submit to the datawarehouse via API
 
 
.DESCRIPTION 
 The webAPI will identify based on your tenantGUID which data sources, and time periods
 are being requested.  This agent will then query the data source(s), collect the 
 data and submit it via the WebAPI for submision to the data warehouse cache database.
 

.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ get-mwp-devices.ps1                                                                         │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 7.18.2019 				               									  │ 
│   AUTHOR      : Paul Drangeid 			                   								  │ 
│   SITE        : https://github.com/pdrangeid/bnwh-cache-agent                               │ 
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#> 

param (
    [string]$subtenant,
    [switch]$noui
    )

$ErrorActionPreference = 'SilentlyContinue'
Remove-Variable -name apikey | Out-Null
Remove-Variable -name tenantguid | Out-Null
$global:srccmdline= $($MyInvocation.MyCommand.Name)
$scriptappname = "Blue Net get-mwp-devices"
$baseapiurl="https://api-cache.bluenetcloud.com"
$ScheduledJobName = "Blue Net Warehouse O365 Refresh"

Write-Host "`nLoading includes: $PSScriptRoot\bg-sharedfunctions.ps1"
Try{. "$PSScriptRoot\bg-sharedfunctions.ps1" | Out-Null}
Catch{
    Write-Warning "I wasn't able to load the sharedfunctions includes (which should live in the same directory as $global:srccmdline). `nWe are going to bail now, sorry 'bout that!"
    Write-Host "Try running them manually, and see what error message is causing this to puke: $PSScriptRoot\bg-sharedfunctions.ps1"
    BREAK
    }

    Prepare-EventLog
    Function Set-CacheSyncJob{

        Get-ScheduledTask -TaskName $ScheduledJobName -ErrorAction SilentlyContinue -OutVariable task |Out-Null
        if ($task -and ![string]::IsNullOrEmpty($subtenant)){
        $tenantjobtaskexists = $false
        Write-Host "Checking Subtentant Task Status"
        $task |
        ForEach-Object {
        if ($_.actions.Arguments -like '*'+$subtenant+'*') {
        # Subtenant already has an action in the existing Scheduled Task
        $tenantjobtaskexists = $true
        }
        if (!$tenantjobtaskexists){
            write-host "This subtenant does not yet have an action item as a part of the scheduled task"
            $answer=yesorno "Would you like to schedule this subtenant refresh job to run automatically?" "Schedule data synchronization"
            if ($answer -eq $true){
            $Username = $env:userdomain+"\"+$Env:USERNAME
            $credentials = $Host.UI.PromptForCredential("Task username and password","Provide the password for this account that will run the scheduled task",$Username,$env:userdomain)
            $Password = $Credentials.GetNetworkCredential().Password 
            $Prog = $env:systemroot + "\system32\WindowsPowerShell\v1.0\powershell.exe"
            $thisuserupn = (get-aduser ($Env:USERNAME)).userprincipalname
            $Opt = '-nologo -noninteractive -noprofile -ExecutionPolicy BYPASS -file "'+$PSScriptRoot+'\get-mwp-devices.ps1" -noui -subtenant '+$subtenant
            $task | ForEach-Object {
                $action = $_.actions
                $action += New-ScheduledTaskAction -Execute $Prog -Argument $Opt -WorkingDirectory $PSScriptRoot
                Set-ScheduledTask -TaskName $ScheduledJobName -Action $action -User $Username -Password $Password
            }# End ForEach-Object (updating tasks)
            }# End User answered YES to adding this task
        }# End subtenantjob action is missing
        }# End ForEach
        }# End have subtenant AND scheduled task
        
        if (!$task) {
        # task does not exist, otherwise $task contains the task object
        $answer=yesorno "Would you like to schedule this agent to run automatically?" "Schedule data synchronization"
        if ($answer -eq $true){
            $Username = $env:userdomain+"\"+$Env:USERNAME
            $credentials = $Host.UI.PromptForCredential("Task username and password","Provide the password for this account that will run the scheduled task",$Username,$env:userdomain)
            $Password = $Credentials.GetNetworkCredential().Password 
            $Prog = $env:systemroot + "\system32\WindowsPowerShell\v1.0\powershell.exe"
            $thisuserupn = (get-aduser ($Env:USERNAME)).userprincipalname
            $Opt = '-nologo -noninteractive -noprofile -ExecutionPolicy BYPASS -file "'+$PSScriptRoot+'\get-mwp-devices.ps1" -noui'
            if (![string]::IsNullOrEmpty($subtenant)){
                $Opt=$Opt+' -subtenant '+$subtenant
            }
            $Action = New-ScheduledTaskAction -Execute $Prog -Argument $Opt  -WorkingDirectory $PSScriptRoot
            $Trigger = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -At "03:30"
            $Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 1 -StartWhenAvailable
            $Settings.ExecutionTimeLimit = "PT10M"
            $Task=Register-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -TaskName $ScheduledJobName -Description "Daily sends updated mwp device data to the reporting datawarehouse via WebAPI" -User $Username -Password $Password -RunLevel Highest
            #$task.triggers.Repetition.Duration ="PT22H"
            #$task.triggers.Repetition.Interval ="PT12M"
            $task | Set-ScheduledTask -User $Username -Password $Password

           
        }# Yes - operater wants us to schedule this task
            }# End if (task doesn't already exist)

        }#End Function
    
    Function submit-cachedata($Cachedata,[string]$DSName){
        write-host "The cache data looks like this `n [$Cachedata]"
    # Takes the resulting cachedata and submits it to the webAPI 
        Write-Host "Submitting Data for $DSName"
        #write-host "******************************* the cache data is: `n"$Cachedata
        $ErrorActionPreference = 'Stop'
        Try{
        $apibase="https://api-cache.bluenetcloud.com/api/v1/submit-data/"
        $apiurlparms="?TenantGUID="+$tenantguid+"&DataSourceName="+$DSName+"&NewTimeStamp="+$querytimestamp
        $apiurl=$apibase+$apiurlparms.replace('+','%2b')
        $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint($apiurl)
        if ($DSName -notlike '*vmware*'){
        #$thecontent = @{"data" = $Cachedata}
        #$thecontent = $($Cachedata | ConvertTo-Json -Depth 5 -Compress)
        #$thecontent = $($Cachedata | ConvertTo-Json -Compress)
        #$thecontent = $(@{"data" = $Cachedata} | ConvertTo-Json -Depth 5 -Compress)
        $thecontent = (@{"data" = $Cachedata} | ConvertTo-Json -Compress)
        }
        $ErrorActionPreference= 'silentlycontinue'
        $pjmb=[math]::Round(([System.Text.Encoding]::UTF8.GetByteCount($Cachedata))*0.00000095367432,2) 
        write-host "Submitting $ic updates for $DSName ($([math]::Round($pjmb,2))MB)"
        if ($DSName -like '*vmware*'){
            $thecontent = $Cachedata
            Invoke-RestMethod $apiurl -Method 'Post' -Headers @{"x-api-key"=$APIKey;"content-type" = "binary"} -Body $thecontent -ErrorVariable RestError -ErrorAction SilentlyContinue -TimeoutSec 900
            }
            else {
        if ($Cachedata -eq "Zero") {
            $thecontent = '{"result":"zero results"}'
        }
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
    
    Function get-mwp-assets([string]$objclass){
        Write-host "getting MWP device assets"
            $ErrorActionPreference = 'Stop'
        
            $o365results=(Get-MsolUser | Select-Object * )
        
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
    if (![string]::IsNullOrEmpty($o365req) ){
        Try{
            write-host "Let's init o365"
            $o365initialized=(get-o365admin $false)
            write-host "got o365admin  and the results are $o365initialized"
        }
        Catch{
            write-host "Sorry - O365 init epic failure!"
        }
        
    }# end initializing O365
   
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
       #Ignoring ADSI requests - this script is O365 only)
    }# end if (ADSI source request)
elseif ($_.SourceName -like "*vmware*"){
       #Ignoring vmware requests - this script is O365 only)
    } # End elseif $_.SourceName -like "*vmware*"
elseif ($_.SourceName -like "o365*"){

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
