function GetIISRoot( [string]$siteName ) {
  $iisWebSite = GetWebsite($siteName)
  new-object System.DirectoryServices.DirectoryEntry("IIS://localhost/" + $iisWebSite.Name + "/Root")
}

function GetVirtualDirectory( [string]$siteName, [string]$vDirName ) {
  $iisWebSite = GetWebsite($siteName)
  $iisVD = "IIS://LocalHost/$($iisWebSite.Name)/ROOT/$vDirName"
  [adsi] $iisVD
}

function DeleteVirtualDirectory( [string]$siteName, [string]$vDirName ) {
  $iisWebSite = GetWebsite($siteName)
  $ws = $iisWebSite.Name
  $objIIS = GetIISRoot $siteName
  write-host "Checking existance of IIS://LocalHost/$ws/ROOT/$vDirName"
  if ([System.DirectoryServices.DirectoryEntry]::Exists("IIS://LocalHost/$ws/ROOT/$vDirName")) {
    write-host "Deleting Virtual Directory $vDirName at $path ..."
    $objIIS.Delete("IIsWebVirtualDir", "$vDirName")
  }
}

function CreateVirtualDirectory( [string]$siteName, [string]$vDirName, [string]$path ) {
  DeleteVirtualDirectory $siteName $vDirName
  $objIIS = GetIISRoot $siteName
  $vDir = $objIIS.psbase.children.add($vDirName,$objIIS.psbase.SchemaClassName)
  $vDir.psbase.CommitChanges()
  $vDir.Path = $path
  $vDir.psbase.CommitChanges()
  $vDir.Put("AuthAnonymous", $true)
  $vDir.Put("AccessExecute", $true)
  $vDir.psbase.CommitChanges()
  write-host "Created Virtual Directory $vdirName at $path"
}

function EnableAllVerbsForSvc([string]$siteName, [string]$vDirName) {
    $vDir = GetVirtualDirectory $siteName $vDirName   
    $vDir.ScriptMaps = $vDir.ScriptMaps -replace "^\.svc.*?$", ".svc,C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\aspnet_isapi.dll,1"
    $vDir.psbase.CommitChanges()
}

function DeleteApplication( [string]$siteName, [string]$vDirName ) {
  write-host "Checking existance of IIS://LocalHost/$ws/ROOT/$vDirName"
  $vd = GetVirtualDirectory $siteName $vDirName
  if ($vd.Path -ne $null) {
    write-host "Deleting Application $vDirName"
    $vd.AppDelete()
  }
}

function CreateApplication( [string]$siteName, [string]$vDirName, [string]$appPoolName ) {
    if ((GetAppPool $appPoolName) -eq $null) {
        throw "Application pool '$appPoolName' does not exist."
    }
  
    $vd = GetVirtualDirectory $siteName $vDirName
    $vd.AppCreate3(2, $appPoolName, 0) #2 = Pooled, 0 = Do not create app pool
    if ($vd.AppGetStatus2() -ne 0) { #0 = hresult 'S_OK... I think.
        throw "Application does not seem to have been created.  Nobody knows why."
    }
    write-host "Created Application $vDirName"
}

function GetWebsite( [string]$siteName ) {
    $iisWebSite = Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class IISWebServerSetting -Filter "ServerComment = '$siteName'"
    if(!$iisWebSite) {
        throw ("No website with the name `"$siteName`" exists on this machine")
    }
    if ($iisWebSite.Count -gt 1) {
        throw ("More than one site with the name `"$siteName`" exists on this machine")
    }
    $iisWebSite
}

function WebsiteExists( [string]$siteName ) {
    (Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class IISWebServerSetting -Filter "ServerComment = '$siteName'") -ne $null
}

function CreateWebsite( [string]$siteName, [int]$port, [string]$path ) {
    $binding = ([wmiclass]"root/MicrosoftIISv2:ServerBinding").CreateInstance()
    $binding.Hostname = ""
    $binding.IP = ""
    $binding.Port = $port
    $iisService = Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class "IIsWebService" -Filter "Name='W3SVC'"
    $iisService.CreateNewSite($siteName, $binding, $path) | Out-Null
    $newSite = (GetWebsite $siteName)
    $server = Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class "IIsWebServer" -Filter "Name='$($newSite.Name)'"
    $server.Start()
}

function IISWeb([string]$task,$application){
  
    $iisweb = "c:\windows\system32\iisweb.vbs"
    switch ($task.toUpper()) 
    { 
        "STOP"  {cscript.exe $iisweb /stop $application} 
        "START" {cscript.exe $iisweb /start $application} 
        
        default {throw "Invalid command: $task"}
    }
}

function RecycleAppPool($appPool){
   cscript.exe "c:\windows\system32\iisapp.vbs" /a $appPool /r      
}

function GetAppPool($name) {
    Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class IIsApplicationPool -Filter "Name = 'W3SVC/AppPools/$name'"
}

function CreateAppPool($appPoolName, $userName, $password) {
    $appPoolSettings = [wmiclass]'root\MicrosoftIISv2:IISApplicationPoolSetting'
    $newPool = $appPoolSettings.CreateInstance()
    $newPool.Name = "W3SVC/AppPools/$appPoolName"
    $newPool.WAMUserName = $userName
    $newPool.WAMUserPass = $password
    $newPool.AppPoolIdentityType = 3
    $result = $newPool.Put()

    $newPool = Get-WmiObject -Namespace 'root\MicrosoftIISv2' -Class IISApplicationPoolSetting -Filter "Name = 'W3SVC/AppPools/$appPoolName'"
    if ($newPool -eq $null) {
        throw "Unable to create ApplicationPool '$appPoolName'"
    }
}

function CreateAppPoolFromPsake($appPoolName) {
    CreateAppPool $appPoolName $script:vickiUserName $script:vickiPassword
}