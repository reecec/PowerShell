$ResultsFile = 'C:\PCINFO\AssetsInfo.csv' # example path for CV results file contaiing pc hostnames specs
$deviceID = 'C:'
$creds = Get-Credential;

function ping-Asset
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $Asset
    )

    Process
    {
       $ADuser = Get-ADUser -Filter "otherHomePhone -like '$Asset'" -Properties otherHomePhone, title, EmailAddress, cn | Select-Object EmailAddress, OtherHomePhone,CN,title
       $ADuserName = $ADuser.cn
       Write-Host ("Asset belongs to: $ADuserName ")
       return $ADuser
    }
}

Function Get-LastLogonDate([string]$AssName) 
{
    # Clean the input
    if ($SAMAccountName -contains $AssName)
    {
        $Where = [array]::IndexOf($SAMAccountName, $AssName)
        $lstLog = $LastLogon[$Where];
        if ($lstLog -eq "01/01/1601")
        {
            return [DateTime]::MinValue;
        }
        else 
        {
            if ($lstLog -as [DateTime])
            {
                return [DateTime]::Parse($lstLog);
            }
            else 
            {
                return [DateTime]::MinValue;
            }
        }
    }

    # If nothing found, return minimum date
    return [DateTime]::MinValue;

} # End of Process



foreach($line in Get-Content "C:\CompSitesPS\hostnames.txt")
{
    Write-Host "********************************"
    Write-Host "Pinging machine $line"
    if (Test-Connection -ComputerName $line -Quiet) 
    {
        Write-Host "pinged asset successfully"  -ForegroundColor Green;
        $pingTEST = ping-Asset $line

        $AssetOwner = $pingTEST.CN
        $AssetOwnerEmail = $pingTEST.EmailAddress
        $AssetOwnerTitle = $pingTEST.title

       Write-Host "$titleTEST" -ForegroundColor Red

        $Results = Get-ADComputer -Identity $line -Properties ObjectSID, ObjectGUID, SamAccountName, CanonicalName, OperatingSystem, OperatingSystemVersion, MemberOf, Distinguishedname, WhenCreated, ipv4Address, ServicePrincipalNames
        $processerInfo =  Get-WmiObject Win32_Processor -ComputerName $line -ErrorAction SilentlyContinue -Credential $creds | Select-Object name;
         $InstallDate =  ([WMI] "").ConvertToDateTime((Get-WmiObject Win32_OperatingSystem -ComputerName $line -Credential $creds).InstallDate)
         
         $disk = Get-WmiObject Win32_logicaldisk -ComputerName $line -Filter "DeviceId='$deviceID'" -Credential $creds| Select-Object DeviceID, freespace, size;
      
        #$office365InstalledOrNot = Get-WmiObject win32_product -ComputerName $line -Credential $creds | where{$_.Name -like "Microsoft Office Professional Plus*"} | Select-Object Name,Version
         
          $VMorNot = Get-WmiObject Win32_ComputerSystem -ComputerName $line -Credential $creds | Select-Object manufacturer, Model, totalphysicalmemory, username;
        
          $patches = get-wmiobject -ComputerName $line -class win32_quickfixengineering -Credential $creds
      
          $drivers = Get-WmiObject -ComputerName $line Win32_PnPSignedDriver -Credential $creds| Select-Object DeviceName, Manufacturer, DriverVersion, signer
      
         [int]$RAM = ($VMorNot.TotalPhysicalMemory/1024/1024/1024);
         [int]$Freespace = ($disk.FreeSpace/1024/1024/1024);
         [int]$Capacity = ($disk.Size/1024/1024/1024);
         $disksInfo = ($disk.DeviceID + " " + $Freespace.ToString() + "/" + $Capacity.ToString());
         $CPU = $processerInfo.Name;
         $driversTest = $drivers.DeviceName
         $LastLogonDate = Get-LastLogonDate $line
         $OS = $Results.OperatingSystem
      
        #Prininting to console
        Write-Host "*****************************************************" -ForegroundColor Green 
        Write-Host "Getting Info for machine: $line" -ForegroundColor Green 
        Write-Host "*****************************************************" -ForegroundColor Green 
        Write-Host "Install date: $test.DateTime"  -ForegroundColor Cyan;
        Write-Host "RAM: $RAM"  -ForegroundColor Cyan;
        Write-Host "CPU: $CPU"  -ForegroundColor Cyan;
        Write-Host "DISKS: $disksInfo"  -ForegroundColor Cyan;
        Write-Host "Results: $OS"  -ForegroundColor Cyan;
        Write-Host "Last Logon for asset was: $LastLogonDate"  -ForegroundColor Cyan;
        Write-Host "OS and security patches are below: $patches  "  -ForegroundColor Cyan;
        Write-Host "drivers installed: $driversTest  "  -ForegroundColor Cyan;
        Write-Host "*****************************************************" -ForegroundColor Green 

        #Custom Objects
        #Asset Info & Asset Owener Info
        New-Object -TypeName PSCustomObject -Property @{
            NAME = $AssetOwner
            EMAIL = $AssetOwnerEmail
            TITLE = $AssetOwnerTitle
            ASSETNUMBER = $line
            INSTALLDATE = $InstallDate
            RAM = $RAM
            CPU = $CPU
            DISKINFO = $disksInfo
            LASTLOGON = $LastLogonDate
            PATCHESTESTTING = (@($patches) -join ',')
            DRIVERS =  (@($driversTest) -join ',')
            } | Export-Csv -Path $ResultsFile -notype
    }   
            else 
            {
                Write-Host "asset did not respond to ping"  -ForegroundColor Red;
            }
}