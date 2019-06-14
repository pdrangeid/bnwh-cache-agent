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

function get-vcentersettings([boolean]$allowpwchange){
    Write-host "Getting vcenter settings..."
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Path = "HKCU:\Software\BNCacheAgent\VMware"
    $ValName = "vCenterServer"	
    $Path = "HKCU:\Software\BNCacheAgent\VMware"
    AddRegPath $Path
    $vCentername = Ver-RegistryValue -RegPath $Path -Name $ValName
        if (AmINull $($vCentername) -eq $true){
    $vCentername="vcsa.domain.local"
    $vCentername = [Microsoft.VisualBasic.Interaction]::InputBox('Enter name of vCenter server.', 'vCenter Server', $($vCentername))
    }
    $vCenterServer=$vCentername.Trim()
        if (AmINull $($vCenterServer) -eq $true){
        write-host "No vCenter Server provided.  Cannot continue this cache collection task..."
        return $false
        }
        Write-Host "Verifying DNS lookup for $vCenterServer"
        Try {
            $ipaddress=$(Resolve-DnsName -Name $vCenterServer -ErrorAction Stop).IPAddress}
        Catch {Write-Warning "Failed to resolve IP address for $vCenterServer.  Check DNS, Firewall, and hostname and try again."
        $Error[0].Exception.Message
        return $false
        }# End Resolve-DnsName failed
        
        Write-Host "Verifying connectivity of vCenter Server at "$ipaddress":"$vcentervalidationport
        $tcpobject = new-Object system.Net.Sockets.TcpClient 
        $connect = $tcpobject.BeginConnect($ipaddress,$vcentervalidationport,$null,$null) 
        $connection = New-Object System.Net.Sockets.TcpClient($ipaddress, $vcentervalidationport)
        $wait = $connect.AsyncWaitHandle.WaitOne(1000,$false) 
        If (-Not $Wait) {
            'Timeout'
            return $PSDefaultParameterValues
        } Else {
            $error.clear()
            $tcpobject.EndConnect($connect) | out-Null 
            If ($Error[0]) {
                Write-warning ("{0}" -f $error[0].Exception.Message)
                return $false
            } Else {
                #'VMware Port open!'
                $vCentername = Ver-RegistryValue -RegPath $Path -Name $ValName -DefValue $vCenterServer
                $ValName = "Passthru"
                $Passthru = Ver-RegistryValue -RegPath $Path -Name $ValName
                if (AmINull $($Passthru) -eq $true){
                $Passthru=YesorNo $("Use passthrough authentication for $vCenterServer"+"?") "Authentication method"
                }# Ask user if we should use passthru
                if ($Passthru -eq $true){
                    write-host "setting passthru to true"
                    Set-ItemProperty -Path $path -Name "Passthru" -Value $Passthru -Force
                }
                if ($Passthru -eq $false){
                    #Now get user/password to authenticate to vCenter Server
                    $ValName = "vCenterUser"
                    $result = Get-Set-Credential $vCenterServer $Path "vCenterUser" "vCenterPW" $false "administrator@vsphere.local"
                    if ($result -eq $true){
                        #Credentials stored
                        Set-ItemProperty -Path $path -Name "Passthru" -Value $Passthru -Force
                    }
                        #Credentials failed to store
                        else {write-host "Failed to store credentials"
                        return $false
                        } #end credential store check
                }#End if (passthru $false)
            }# End Else (no error with vmware port)
             }# Else done waiting
             #Now check for RVTools
             [string] $RVToolsPath = ${env:ProgramFiles(x86)}+"\Robware\RVTools"
             [string] $global:RVToolsPathExe = ${RVToolsPath}+"\RVTools.exe"
             if(![System.IO.File]::Exists($RVToolsPathExe)){
                # RVTools executable doesn't seem to exist
                Write-Warning "This workstation does not have rvtools installed.  Please download and install and re-run the script"
                Start "https://www.robware.net/rvtools/download/"
                return $false
            }
        return $true
}#End Function (get-vcentersettings)

Function get-vmware-assets([string]$objclass){
Write-host "getting vmware assets"
    $ErrorActionPreference = 'Stop'
# -----------------------------------------------------
# Set parameters for vCenter and start RVTools export
# -----------------------------------------------------
$Path = "HKCU:\Software\BNCacheAgent\VMware"
$vCentername=Ver-RegistryValue -RegPath $Path -Name "vCenterServer"
$Passthru=Ver-RegistryValue -RegPath $Path -Name "Passthru"
$vmwarecsv=New-TemporaryDirectory

#[string] $VCServer = $(Ver-RegistryValue -RegPath $Path -Name $ValName)
write-host "passthru is $Passthru and vcenter is $vCentername.  let's proc class:$objclass"
if ($passthru -eq $true){
    Write-Host "Using Passthru authentication"
    $objname = "Export"+$objclass+"2csv"
    $Arguments = " -passthroughAuth -s $vCentername -c $objname -d $($vmwarecsv)"
    Write-Host "Args: $Arguments"
    $Process = Start-Process -FilePath $RVToolsPathExe -ArgumentList $Arguments -NoNewWindow -Wait
}

if ($passthru -eq $false){
    Write-Host "Using User-supplied credentials"
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
Write-Host "Export Success!"
Write-host "Now do some processing and uploading..."
write-host "returning the file ($vmwarecsv) for processing"
return  $($vmwarecsv)
}
