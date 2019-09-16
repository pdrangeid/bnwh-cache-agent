<# 
.SYNOPSIS 
 Functions to configure and query VMware vCenter using rvtools and then submit the resulting data files
 to the warehouse cache via webAPI
 
.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ get-vmware-data.ps1                                                                         │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 5.28.2019 				               									  │ 
│   AUTHOR      : Paul Drangeid 			                   								  │ 
│   SITE        : https://github.com/pdrangeid/bnwh-cache-agent                               │ 
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#> 
# This is the port to validate that a vCenter server is running at the destination.
$vcentervalidationport=9443
$esxivalidationport=902

Function get-portvalidation([string]$hostorip,[int]$tcpport){
    $ErrorActionPreference = 'Silently Continue'
    $tcpobject = new-Object system.Net.Sockets.TcpClient 
    #$connect = $tcpobject.BeginConnect($hostorip,$tcpport,$null,$null) 
    #$wait = $connect.AsyncWaitHandle.WaitOne(1000,$false) 
    
    #$connection = New-Object System.Net.Sockets.TcpClient($hostorip, $tcpport)
    $connect = $tcpobject.BeginConnect($hostorip,$tcpport,$null,$null) 
    $wait = $connect.AsyncWaitHandle.WaitOne(1000,$false) 
    If (-Not $Wait) {
        #'Timeout'
        return $false
    } Else {
        $error.clear()
        $tcpobject.EndConnect($connect) | out-Null 
        If ($Error[0]) {
            return $false
        } Else {
            # port responded!
            return $true
        } # Port responded
    } #End if error

}# End Function get-portvalidation

function get-vcentersettings([switch]$allowpwchange){
    $ErrorActionPreference = 'Silently Continue'
    Show-onscreen $("Getting vcenter settings...") 2
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Path = "HKCU:\Software\BNCacheAgent\VMware"
    $ValName = "vCenterServer"	
    $Path = "HKCU:\Software\BNCacheAgent\VMware"
    AddRegPath $Path
    $vCentername = Ver-RegistryValue -RegPath $Path -Name $ValName
            if ([string]::IsNullOrEmpty($vCentername) -or $vCentername -eq ''){
    if ($noui -eq $true){
        Write-Warning "VMware data was requested, but vCenter/ESXi has not been configured, and we are running as a scheduled task.  Terminating VMware module."
        return $false
    }
    $vCentername="vcsa.domain.local"
    $vCentername = [Microsoft.VisualBasic.Interaction]::InputBox('Enter name of vCenter server (or standalone ESXi host).', 'vCenter Server', $($vCentername))
    }
        if (AmINull $($vCentername) -eq $true){
        write-host "No Server provided.  Cannot continue this cache collection task..."
        return $false
        }
        $vCenterServer=$vCentername.trim()
        Show-onscreen $("Verifying DNS lookup for $vCenterServer") 1
        Try {
            $ipaddress=$(Resolve-DnsName -Name $vCenterServer -ErrorAction Stop).IPAddress}
        Catch {Write-Warning "Failed to resolve IP address for $vCenterServer.  Check DNS, Firewall, and hostname and try again."
        $Error[0].Exception.Message
        return $false
        }# End Resolve-DnsName failed
        
        $vmode = Ver-RegistryValue -RegPath $Path -Name "Mode"
        if ($vmode -eq 'standalone'){
            $queryport = $esxivalidationport
        } else {$queryport = $vcentervalidationport}
        Show-onscreen $("Verifying connectivity of vCenter Server at "+$ipaddress+":"+$queryport) 1
        $result =  get-portvalidation $ipaddress $queryport
        if ($result -eq $false){
            $Tryesxi=YesorNo $("No response when querying for vCenterServer Would you like to try querying as a standalone ESXi"+"?") "No vCenter response... Try standalone host?"
            if ($Tryesxi -eq $true){
            $result =  get-portvalidation $ipaddress $esxivalidationport
            if ($result -eq $true){
                Set-ItemProperty -Path $path -Name "Mode" -Value "standalone" -Force
            }
            } #Try as standalone ESXi Host?
        }# vCenter port validation attempt failed
        if ($result -ne $true){
              return $result
        }
                #'VMware Port open!'
                $vCentername = Ver-RegistryValue -RegPath $Path -Name $ValName -DefValue $vCenterServer
                $ValName = "Passthru"
                $Passthru = Ver-RegistryValue -RegPath $Path -Name $ValName
                if ([string]::IsNullOrEmpty($Passthru)){
                $Passthru=YesorNo $("Use passthrough authentication for $vCenterServer"+"?") "Authentication method"
                }# Ask user if we should use passthru
                if ($Passthru -eq $true){
                    Show-onscreen $("VMware Passthru authentication: $Passthru") 2
                    Set-ItemProperty -Path $path -Name "Passthru" -Value $Passthru -Force
                }
                if ($Passthru -eq $false){
                    #Now get user/password to authenticate to vCenter Server
                    $ValName = "vCenterUser"
                    $result = Get-Set-Credential $vCenterServer $Path "vCenterUser" "vCenterPW" $false "administrator@vsphere.local"
                    if ($result -eq $true){
                        #Credentials stored
                        Set-ItemProperty -Path $path -Name "Passthru" -Value $Passthru -Force
                    }# Got creds
                        else 
                        {#Credentials failed to store
                        write-host "Failed to store credentials"
                        return $false
                        } #end credential store check
                }#End if (passthru $false)
            #}# End Else (no error with vmware port)
             
             #Now check for RVTools
             [string] $RVToolsPath = ${env:ProgramFiles(x86)}+"\Robware\RVTools"
             [string] $global:RVToolsPathExe = ${RVToolsPath}+"\RVTools.exe"
             if(![System.IO.File]::Exists($RVToolsPathExe)){
                # RVTools executable doesn't seem to exist
                Write-Warning "This workstation does not have rvtools installed.  Please download and install and re-run the script"
                Start-Process "https://www.robware.net/rvtools/download/"
                return $false
            }
            $global:RVToolsVersion=[System.Diagnostics.FileVersionInfo]::GetVersionInfo($global:RVToolsPathExe).FileVersion
            if ($global:RVToolsVersion -lt "3.11.6.0.0"){Write-Warning "Found RVTools (version $global:RVToolsVersion), but this version has known issues.  Please update."}
        return $true
}#End Function (get-vcentersettings)

Function get-vmware-assets([string]$objclass){
    Show-onscreen $("Querying VMware assets") 2
    $ErrorActionPreference = 'Stop'
# -----------------------------------------------------
# Set parameters for vCenter and start RVTools export
# -----------------------------------------------------
if ($global:RVToolsVersion -lt "3.11.6.0.0"){
    write-warning "RVToolsVersion is older, and has known issues with some object class exports.  Please upgrade."
    if($objclass -eq 'Dvport' -or $objclass -eq 'vSC+VMK'){
    write-warning "This version of RVTools does not support exporting $objclass.  Please upgrade."
    return $false
    }# Request for VMware data that is broken with the provided version of rvtools
}# Detected your version of rvtools is not at least 3.11.6
$Path = "HKCU:\Software\BNCacheAgent\VMware"
$vCentername=Ver-RegistryValue -RegPath $Path -Name "vCenterServer"
$Passthru=Ver-RegistryValue -RegPath $Path -Name "Passthru"
$vmwarecsv=New-TemporaryDirectory

#[string] $VCServer = $(Ver-RegistryValue -RegPath $Path -Name $ValName)
Show-onscreen $("VMware passthru authentication: $Passthru `nvCenter target: $vCentername `Querying object class:$objclass") 2
if ($passthru -eq $true){
    Write-Host "Using Passthru authentication"
    $objname = "Export"+$objclass+"2csv"
    $Arguments = " -passthroughAuth -s $vCentername -c $objname -d $($vmwarecsv)"
    Write-Host "Args: $Arguments"
    $Process = Start-Process -FilePath $RVToolsPathExe -ArgumentList $Arguments -NoNewWindow -Wait
}

if ($passthru -eq $false){
    Show-onscreen $("Querying $vCentername with User-supplied credentials") 2
    $vcuser=Ver-RegistryValue -RegPath $Path -Name "vCenterUser"
    $vcuserpw=Get-SecurePassword $Path "vCenterPW"
    $objname = "Export"+$objclass+"2csv"
    $Arguments = " -u $vcuser -p $vcuserpw -s $vCentername -c $objname -d $($vmwarecsv)"
    $Process = Start-Process -FilePath $RVToolsPathExe -ArgumentList $Arguments -NoNewWindow -Wait
    Remove-Variable -name vcuserpw | Out-Null
    Remove-Variable -name vcuser | Out-Null
}
if($Process.ExitCode -eq -1)
{
    Write-Host "Error: RVTools Export failed! RVTools returned exitcode -1" -ForegroundColor Red
    return $false
}
Show-onscreen $("returning the file ($vmwarecsv) for processing") 2
return  $($vmwarecsv)
}
