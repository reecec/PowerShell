$WindowsServers =  "SERVER"

#loop through each index in array containing computer object name
foreach($WindowsServer in $WindowsServers)
{
#if computer responds to ping request within 10 seconds do the following
$pingtest = test-connection -computername $WindowsServer -Quiet -Count 10

#if ping test return boolean true then do the following
if ($pingtest -match "True") {
#check to see if winRM enabled for machine
$checkWinRM =  Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)"

if($checkWinRM.Enabled -match "False")
{
    Enable-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)"
}

#check to see if IIS service is on machine and store get-service object in IIS_WWPS variable
$IIS_WWWPS = get-service -computername $WindowsServer | Where-Object {$_.DisplayName -eq "World Wide Web Publishing Service"} | select machineName, Name, DisplayName, Status
#check to see if WAS service (responsible for creation of worker processers/app pools) is on machine and store get-service object in IIS_WAS variable
$IIS_WAS = get-service -computername $WindowsServer | Where-Object {$_.DisplayName -eq "Windows Process Activation Service"} | select machineName, Name, DisplayName, Status
#check to see if SQL service is on machine and store get-service object in SQL_SERVICE variable
$SQL_SERVICE = get-service -computername $WindowsServer | Where-Object {$_.DisplayName -contains "*SQL*"} | select machineName, Name, DisplayName, Status

    #if the IIS_WWPS objects property displayname is WWWPS and SQL_SERVICE objects property displayname Contains SQL do the following
    if($IIS_WWWPS.DisplayName -match "World Wide Web Publishing Service" -And $IIS_WAS.DisplayName -match "Windows Process Activation Service" -And $SQL_service.DisplayName -contains "*SQL*")
    {
    Write-Output "Both IIS and SQL installed on machine $WindowsServer" 
    }
    #if both IIS variable objects, displayname property matches IIS service names do the following
    elseif($IIS_WWWPS.DisplayName -match "World Wide Web Publishing Service" -And $IIS_WAS.DisplayName -match "Windows Process Activation Service")
    {
    Write-Output "IIS is installed on machine $WindowsServer" 
    
    #uisng invoke-commandfunction to run script on remote machine and containing cmdlets in script block
    Invoke-Command -ComputerName $WindowsServer -ScriptBlock { 
    
    #check registry of remote machine to return version of IIS via keys string value (-a account required)
    get-itemproperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\InetStp\ | Select PSComputerName, IISProgramGroup, SetupString
    
    #importing webAdmin module for IIS cmdlets to return websites and app pools on server
    Import-Module WebAdministration; iis:; $websites = Get-Website | Select applicationPool, PScomputerName, Name, PhysicalPath, bindings; $websites; foreach($websitebindings in $websites){get-webBinding -name $websitebindings.Name}} #-Credential corp\Rcunni10-a (USE IF NEED EVLAUATED IE -AD)
    }

    elseif($SQL_SERVICE -contains "*SQL*")
    {
     Write-Output "SQL is installed on machine $WindowsServer"
     #*****************************************************************************************
     # FIND SQL REGISTRY PATH AND KEY STRING VLAUES FOR DIFFERENT VERSIONS AND ADD SWITCH-CASE*
     #*****************************************************************************************
     get-itemproperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\ | Select PSComputerName, IISProgramGroup, SetupString
    }
    #when neither SQL or IIS instance found on machine
    else
    {
    Write-Output "Neither SQL or IIS is installed on machine $WindowsServer"  
    }

}
else
    {
      Write-Output "$WindowsServer did not respond to ping request"
    }
}