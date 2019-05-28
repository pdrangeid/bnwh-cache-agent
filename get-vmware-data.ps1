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
function get-vcentersettings(){
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
        BREAK
        }
    
        Write-Host "Verifying DNS lookup for $vCenterServer"
        Try {
            $ipaddress=$(Resolve-DnsName -Name $vCenterServer -ErrorAction Stop).IPAddress}
        Catch {Write-Warning "Failed to resolve IP address for $vCenterServer.  Check DNS, Firewall, and hostname and try again."
        $Error[0].Exception.Message
        exit
        }# End Resolve-DnsName failed
        
        Write-Host "Verifying connectivity of vCenter Server at "$ipaddress":"$vcentervalidationport
        $tcpobject = new-Object system.Net.Sockets.TcpClient 
        $connect = $tcpobject.BeginConnect($ipaddress,$vcentervalidationport,$null,$null) 
        $connection = New-Object System.Net.Sockets.TcpClient($ipaddress, $vcentervalidationport)
        $wait = $connect.AsyncWaitHandle.WaitOne(1000,$false) 
        If (-Not $Wait) {
            'Timeout'
        } Else {
            $error.clear()
            $tcpobject.EndConnect($connect) | out-Null 
            If ($Error[0]) {
                Write-warning ("{0}" -f $error[0].Exception.Message)
            } Else {
                #'VMware Port open!'
                $vCentername = Ver-RegistryValue -RegPath $Path -Name $ValName -DefValue $vCenterServer
                $ValName = "Passthru"
                $Passthru = Ver-RegistryValue -RegPath $Path -Name $ValName
                if (AmINull $($Passthru) -eq $true){
                $Passthru=YesorNo "Should we use passthrough authentication for querying $vCenterServer?" "Authentication method"
                }# Ask user if we should use passthru
                if ($Passthru -eq $false){
                    #Now get user/password to authenticate to vCenter Server
                    $ValName = "vCenterUser"
                    if ((Test-RegistryValue -Path $Path -Value $ValName) -and (Test-RegistryValue -Path $Path -Value "vCenterPW")){
                        write-host "Oh good! -- It looks like we already have a validated user/pw stored!"
                        write-host "Would you like to update the credentials used for [$vCentername]?"
                        $intsetacct=-1
                        if (YesorNo "Would you like to update the credentials used for [$vCentername]?" "vCenter Credentials") {
                        $intsetacct=1
                        }
                        } # End if (username AND password stored in registry)
                }
            }
     
        }
    }
