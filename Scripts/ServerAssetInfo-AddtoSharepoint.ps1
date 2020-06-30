# =====================================================================================
# Get data for the servers on the estate, add to sharepoint list, Placement 2017
# =====================================================================================
# Reece Cunningham/Darren Quinn:20170715
# =====================================================================================

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

# User varilables
# Header for the output file.  Leave blank if nothing special about the run.
$HeaderName = "";
$Environment = "Production";
$InputFilename = "\Assets.csv";
$creds = Get-Credential;

# Specify tenant admin and site URL
$SiteUrl = "URL of site"
$ListName = "WindowsServers"
$UserName = "UPN"
$SecurePassword = ConvertTo-SecureString "EnterPassword" -AsPlainText -Force

# Create Timestamp for output files
$FileTimeStamp = $(Get-Date -Format "yyyyMMdd-HHmmss");

# Initialise file and folder variables
$WorkingFullpath = (Get-Variable MyInvocation).Value;
$WorkingFolder = Split-Path $WorkingFullpath.MyCommand.Path;
$TempFolder = $WorkingFolder + "\Temp\";
$OutputFilePrefix = $WorkingFolder + "\Output\AssetInfo";

# Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path “C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll”
Add-Type -Path “C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll”
Add-Type -Path “C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll”

# Bind to site collection
$ClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$ClientContext.Credentials = $credentials
$ClientContext.ExecuteQuery()

# Get List
$List = $ClientContext.Web.Lists.GetByTitle($ListName)
$ClientContext.Load($List)
$ClientContext.ExecuteQuery()

# Create output files| pinged success and ping failed.
$InputFilename = ($WorkingFolder + $InputFilename);
$ResultOutputFile = $OutputFilePrefix + $HeaderName + $FileTimeStamp + ".csv";
$ResultFailedOutputFile = $OutputFilePrefix + $HeaderName + "_Failed_" + $HeaderName + $FileTimeStamp + ".csv";
Out-File $ResultOutputFile -InputObject "ASSETNAME,PING,LASTLOGONDATE,TYPE,RAM,MANUFACTURER,MODEL,CPU,DISKS,LastUser,OS,OSVERSION,ServicePack,ROLES,DN,WHEN CREATED,IPV4,GUID,CN,ENVIRONMENT,DATACOLLECTIONCOMMENT";
Out-File $ResultFailedOutputFile -InputObject "ASSETNAME,PING,LASTLOGONDATE,TYPE,RAM,MANUFACTURER,MODEL,CPU,DISKS,LastUser,OS,OSVERSION,ServicePack,ROLES,DN,WHEN CREATED,IPV4,GUID,CN,ENVIRONMENT,DATACOLLECTIONCOMMENT";
Out-File $InputFilename -InputObject "AssetName";

# Start time for use to see how long it ran for
$StartTime = [DateTime]::Now; 

Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host "Starting Run"                 -ForegroundColor Cyan;
Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host ("Start Time    : " + $StartTime)  -ForegroundColor Cyan;
Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host "";

# Get the latest last logon file
Write-Host (" - Get newest filename") -ForegroundColor Green;
$LastLoggedOnFolder = "\\servername\d$\_ADExports_Computers\";
$latest = Get-ChildItem -Path $LastLoggedOnFolder | Sort-Object LastAccessTime -Descending | Select-Object -First 1;
Write-Host (" - Copy file " + $latest.FullName) -ForegroundColor green;
Copy-Item $latest.FullName ($TempFolder + "LastLogon.csv");

# Load the file into memory (Just the servers)
$LogonFile = New-Object IO.StreamReader ($TempFolder + "LastLogon.csv");
$row = $LogonFile.ReadLine();
$SAMAccountName = @()
$LastLogon = @()

# Loop through each row loading just the servers and store it in an array
While ($row -ne $null)
{ 
    $Fields = $row.ToString().Split(",");
    if ($Fields[2] -ne $null) 
    {
        if ($Fields[2].ToString().Length -ne 0)
        {
            if ($Fields[2].ToUpper() -like "*SERVER*")
            {
                $SAMAccountName += $Fields[0];
                $LastLogOn += $Fields[1];
                Out-File $InputFilename -InputObject $Fields[0] -Append;
            }
        }
    }
    $row = $LogonFile.ReadLine();
}
$LogonFile.Close();

# Open input file
$ServersCount = Import-Csv $InputFilename | Measure-Object | Select Count
$Servers = Import-Csv $InputFilename;
$RowCount = 0;
$UpdateCnt = 0;
$InsertCnt = 0;
$ErrCnt = 0;
$pingConfig = 
@{ 
    "count" = 1
    "bufferSize" = 15
    "delay" = 1
    "EA" = 0 
}

# Loop through the server list
foreach($Server in $Servers) 
{
    $RowCount += 1;
    Write-Host ("Getting data for asset: " + $Server.AssetName + " (" + $rowcount.ToString() + "/" + $ServersCount.Count.ToString() + ")") -ForegroundColor Yellow
    $Results = Get-ADComputer -Identity $Server.AssetName -Server servername -Properties ObjectSID, ObjectGUID, SamAccountName, CanonicalName, OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, MemberOf, Distinguishedname, WhenCreated, ipv4Address, ServicePrincipalNames
	if (Test-Connection -ComputerName $Server.AssetName @pingConfig -Quiet) 
    {
        $pingDate = [DateTime]::Now.Date;
        Write-Host (" - Pinged successfully @ " + $pingDate) -ForegroundColor Green

        # Get the WMI Calls, sometimes they fail with access denied, stop doing them if they do
        try
        {
		    # Make all the WMI calls in one place
            $Error.Clear();
            Write-Host (" - Getting WMI Data...") -ForegroundColor Green
            $VMorNot = Get-WmiObject Win32_ComputerSystem -ComputerName $Server.AssetName -Credential $creds -ErrorAction SilentlyContinue | select manufacturer, Model, totalphysicalmemory, username;
            if ($Error.Count -eq 0)
            {
		        $processerInfo =  Get-WmiObject Win32_Processor -ComputerName $Server.AssetName -Credential $creds -ErrorAction SilentlyContinue | select name;
		        $disks = Get-WmiObject  Win32_logicaldisk -ComputerName $Server.AssetName -Credential $creds -ErrorAction SilentlyContinue | select DeviceID, freespace, size;
                Write-Host (" - Completed WMI Data retrieval") -ForegroundColor Green

                # Is the server virtual or physical
                if($VMorNot.manufacturer -like "*VM*" -or $VMorNot.Model -like "*VM*")
		        { $PCtype = "Virtual Machine"; }
		        else { $PCtype = "Physical Machine"; }

                # Get the amount of RAM in GB
		        [int]$RAM = ($VMorNot.TotalPhysicalMemory/1024/1024/1024);
		    
                # Get some other information
                $Manufacturer = $VMorNot.Manufacturer;
		        $Model = $VMorNot.Model;
		        $CPU = ($processerInfo.Count.ToString() + " x " + $processerInfo.Name[0]);
		        $LastUser = $VMorNot.Username;
		    
                # Get the LUN info
		        $disksInfo = "";
		        foreach($disk in $disks)
		        {
                    # Ignore floppy disk
                    if ($disk.DeviceID -ne "A:")
                    {
                        # Round off the disk sizes
                        [int]$Freespace = ($disk.FreeSpace/1024/1024/1024);
                        [int]$Capacity = ($disk.Size/1024/1024/1024);
                        if ($disksInfo.Length -ne 0) { $disksInfo = $disksInfo  + ","; }
    		            $disksInfo = ($disksInfo + $disk.DeviceID + " " + $Freespace.ToString() + "/" + $Capacity.ToString());
                    }
		        }

                $DataCollectionComment = "Good";
                Write-Host (" - Finished data collection successfully") -ForegroundColor Green
            }
            else 
            {
                $DataCollectionComment = ("An error occured with WMI. " + $Error[0].Exception.Message);
                Write-Host (" - $DataCollectionComment") -ForegroundColor Magenta
                $Manufacturer = "";
                $Model = "";
                $RAM = "";
                $LastUser = "";
                $disksInfo = "";
                $PCtype = "";
            }
		}
        catch
        {
            $DataCollectionComment = ("WMI failed to connect. " + $_.Exception.Message);
            Write-Host (" - $DataCollectionComment") -ForegroundColor Magenta
            $Manufacturer = "";
            $Model = "";
            $RAM = "";
            $LastUser = "";
            $disksInfo = "";
            $PCtype = "";
        }

        $LastLogonDate = Get-LastLogonDate $Server.AssetName;
        $RolesInstalled = "";

        Out-File $ResultOutputFile -InputObject($Server.AssetName + "," + "," + $pingDate + "," + $LastLogonDate + "," + $PCtype + "," + $RAM + ",""" + $Manufacturer + """,""" + $Model + """,""" + $CPU + """,""" + $disksInfo + """," + $LastUser + ",""" + $Results.OperatingSystem + """," + $Results.OperatingSystemVersion + "," + $Results.OperatingSystemServicePack + "," + $RolesInstalled + ",""" + $Results.DistinguishedName + """," + $Results.WhenCreated + "," + $Results.IPv4Address + "," + $Results.ObjectGUID + ",""" + $Results.CanonicalName + """," + $Environment + ",""" + $DataCollectionComment + """") -Append;
    }
	else
	{
        $DataCollectionComment = (" - Ping failed.");
        Write-Host (" - Ping failed") -ForegroundColor Magenta
        $Manufacturer = "";
        $Model = "";
        $RAM = "";
        $LastUser = "";
        $disksInfo = "";
        $PCtype = "";
        Out-File $ResultFailedOutputFile -InputObject($Server.AssetName + "," + $pingDate + "," + $LastLogonDate + "," + $PCtype + "," + $RAM + ",""" + $Manufacturer + """,""" + $Model + """,""" + $CPU + """,""" + $disksInfo + """," + $LastUser + ",""" + $Results.OperatingSystem + """," + $Results.OperatingSystemVersion + "," + $Results.OperatingSystemServicePack + "," + $RolesInstalled + ",""" + $Results.DistinguishedName + """," + $Results.WhenCreated + "," + $Results.IPv4Address + "," + $Results.ObjectGUID + ",""" + $Results.CanonicalName + """," + $Environment + ",""" + $DataCollectionComment + """") -Append;
	}	

    # Search for list item
    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery;
    $Query.ViewXml = "<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>" + $Server.AssetName + "</Value></Contains></Where></Query></View>";
    $ListItem = $list.GetItems($query);
    $ClientContext.Load($ListItem);
    $ClientContext.ExecuteQuery();

    # Update if found, insert if not
    if ($ListItem.Count -eq 1)
    {
        # Update ListItem
        $ListItem[0]["Guid0"] = $Results.ObjectGUID;
        $ListItem[0]["_x0049_PV4"] = $Results.IPv4Address;
        $ListItem[0]["LastSuccessfulPing"] = $pingDate; #.ToString("dd/MM/yyyy");
        if ($LastLogonDate -ne [DAteTime]::MinValue) { $ListItem[0]["LastLogonDate"] = $LastLogonDate; }
        $ListItem[0]["ServerType"] = $PCtype;
        $ListItem[0]["RAM"] = $RAM;
        $ListItem[0]["MachineManufacturer"] = $Manufacturer;
        $ListItem[0]["ServerModel"] = $Model;
        $ListItem[0]["CPU"] = $CPU;
        $ListItem[0]["Disks"] = $disksInfo;
        $ListItem[0]["LastAccessedBy"] = $LastUser;
        $ListItem[0]["OperatingSystem"] = $Results.OperatingSystem;
        $ListItem[0]["OSVersion"] = $Results.OperatingSystemVersion;
        $ListItem[0]["ServicePack"] = $Results.OperatingSystemServicePack;
        $ListItem[0]["DistinguishedName"] = $Results.DistinguishedName;
        $ListItem[0]["ServerCreatedDate"] = $Results.WhenCreated; #.ToString("dd/MM/yyyy");
        $ListItem[0]["CN"] = $rec.CN;
        $ListItem[0]["Environment"] = $Environment;
        $ListItem[0]["DataCollectionComment"] = $DataCollectionComment;

        $ListItem.Update();
        $ClientContext.ExecuteQuery();
        $UpdateCnt++;
    }
    else
    {
        $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $NewListItem = $List.AddItem($ListItemCreationInformation)
        $NewListItem["Title"] = $Server.AssetName;
        $NewListItem["AssetName"] = $Server.AssetName;
        $NewListItem["Guid0"] = $Results.ObjectGUID;
        $NewListItem["_x0049_PV4"] = $Results.IPv4Address;
        $NewListItem["LastSuccessfulPing"] = $pingDate; #.ToString("dd/MM/yyyy");
        if ($LastLogonDate -ne [DAteTime]::MinValue) { $NewListItem["LastLogonDate"] = $LastLogonDate; }
        $NewListItem["ServerType"] = $PCtype;
        $NewListItem["RAM"] = $RAM;
        $NewListItem["MachineManufacturer"] = $Manufacturer;
        $NewListItem["ServerModel"] = $Model;
        $NewListItem["CPU"] = $CPU;
        $NewListItem["Disks"] = $disksInfo;
        $NewListItem["LastAccessedBy"] = $LastUser;
        $NewListItem["OperatingSystem"] = $Results.OperatingSystemVersion;
        $NewListItem["OSVersion"] = $Results.OperatingSystemVersion;
        $NewListItem["ServicePack"] = $Results.OperatingSystemServicePack;
        $NewListItem["DistinguishedName"] = $Results.DistinguishedName;
        $NewListItem["ServerCreatedDate"] = $Results.WhenCreated; #.ToString("dd/MM/yyyy");
        $NewListItem["CN"] = $rec.CN;
        $NewListItem["Environment"] = $Environment;
        $NewListItem["DataCollectionComment"] = $DataCollectionComment;

        $NewListItem.Update();
        $ClientContext.ExecuteQuery();
        $InsertCnt++;
    }
}

$TimeTaken = ([DateTime]::Now.Subtract($StartTime));
$EndTime = [DateTime]::Now;

Write-Host "";
Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host "Start Time    : $StartTime  "  -ForegroundColor Cyan;
Write-Host "End Time      : $EndTime"      -ForegroundColor Cyan;
Write-Host "Time Taken    : $TimeTaken  "  -ForegroundColor Cyan;
Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host "New Records   : $InsertCnt  "  -ForegroundColor Cyan;
Write-Host "Updated       : $UpdateCnt  "  -ForegroundColor Cyan;
Write-Host "Errors        : $ErrCnt     "  -ForegroundColor Cyan;
Write-Host "==============================================" -ForegroundColor Cyan;
Write-Host ("Output File   : " + $ResultOutputFile) -ForegroundColor Cyan;
Write-Host ("Error File    : " + $ResultFailedOutputFile) -ForegroundColor Cyan;
Write-Host "==============================================" -ForegroundColor Cyan;
