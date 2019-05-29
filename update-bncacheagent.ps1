<# 
.SYNOPSIS 
 Update the agent with the latest version from the repository
 
 
.DESCRIPTION 
 Setup and maintain agent scripts with latest version from the git repository.
 

.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ update-bncacheagent.ps1                                                                     │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 5.28.2019 				               		     			    │ 
│   AUTHOR      : Paul Drangeid 			                   				    │ 
│   SITE        : https://github.com/pdrangeid/bnwh-cache-agent                               │ 
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#> 

$companyname="Blue Net Inc"
$reporoot="https://raw.githubusercontent.com/pdrangeid"
$path = $("$Env:Programfiles\$companyname")
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

$path = $("$Env:Programfiles\$companyname\Caching Agent")
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

$client = new-object System.Net.WebClient
$client.DownloadFile("$reporoot/n4j-pswrapper/master/bg-sharedfunctions.ps1","$path\bg-sharedfunctions.ps1")
$client.DownloadFile("$reporoot/bnwh-cache-agent/master/get-datawarehouse-cache.ps1","$path\get-datawarehouse-cache.ps1")
$client.DownloadFile("$reporoot/bnwh-cache-agent/master/get-vmware-data.ps1","$path\get-vmware-data.ps1")
$client.DownloadFile("$reporoot/bnwh-cache-agent/master/update-bncacheagent.ps1","$path\update-bncacheagent.ps1")


$env:Path += ";$path"
