<#
Server Inventory
Daryl Bizsley 2014

Edit the following Parameters to your specifications

$Logo this can be your company logo, this is set to a windows logo by default (Filepath is relative to resource directory($resourcedir))
$filepath this is the directory to output the files to.
$enableschedtask this can have values of true or false depending on if you would like a scheduled task to run this script. (It will run daily at 4.00AM by default)
$Scheduledtaskuser this sets the username you would like to run the scheduled task as, must have logon  as a batch job priviledges
$Scheduledtaskpass this set the password of the user for scheduled task
$Scriptdir this is the location to copy the script to on the machine for running as a scheduled task
$sendemail set to true if you want generated emails (No SMTP Authentication is implemented)
$MasterScriptVersion Set this location so that script can auto update on all servers/clients if you need to add anything
#>


#$ErrorActionPreference = "SilentlyContinue" #Some non fatal errors may occur uncomment this line if you don't want to display them.
######Variables to fill
$MasterScriptVersion = "\\Server\share\ServerInventory.ps1" 
$logo = "Logo.png"
$filepath = "\\someserver\someshare"
$resourcedir =  "./Resources"
$enableschedtask = $true
$Scheduledtaskuser = "domainname\username"
$Scheduledtaskpass = "PlaintextPassword"
$Scriptdir = "C:\ServerInventory\"
$sendemail = $false



#######
## Email Setting if you would like to email the Results
###Edit as appropriate
$smtp = "Your-ExchangeServer"
$to = "YourIT@YourDomain.com"
$subject = "Hardware Info of $name"
$attachment = "$filepath\$name.html"
$from =  (Get-Item env:\username).Value + "@yourdomain.com"


#####Some needed Variables to populate
$name = (Get-Item env:\Computername).Value
$description = Get-WmiObject -Class Win32_OperatingSystem |Select Description
$description = $description.description
$date = Get-Date




#### HTML Output Formatting #######
$Virtual = (Get-WmiObject win32_systemenclosure -Property Manufacturer)
$CPU = (Get-WmiObject win32_processor -Property Manufacturer)
###Images at top
if ($sendemail -ne $true){
    $a = "<link rel='stylesheet' type='text/css' href='$resourcedir/WEB/theme.css'>"
    

    $a = $a + "<img src ='$resourcedir/WEB/$logo' alt='logo'>"
    if ($Virtual.Manufacturer -eq "No Enclosure"){
        $a = $a + "<img src ='$resourcedir/WEB/VMWare.png' alt='vmware'>"
    }
    Elseif ($CPU.Manufacturer -eq "GenuineIntel")  {
        $a = $a + "<img src='$Resourcedir/WEB/PhysicalIntel.png' alt='intel'>"
    }
    Else {
        $a = $a + "<img src='$resourcedir/WEB/PhysicalAMD.png' alt='AMD'>"
    }
}
Else {
    $css = Get-Content "$filepath\$resourcedir\WEB\theme.css"
    $a = "<style>$css</style>"
    } 
#### Setting up header and description
ConvertTo-Html -Head $a  -Title "Hardware Information for $name" >  "$filepath\$name.html"
ConvertTo-Html -Body "<br><br><br><h1>$name</h1>" >> "$filepath\$name.html"
ConvertTo-Html -Body "<description>$description</Description>" >> "$filepath\$name.html"
###Check if physical or VM###
if ($Virtual.Manufacturer -eq "No Enclosure"){
    ConvertTo-Html -Body "<servertype>Virtual Server</servertype>" >> "$filepath\$name.html"
}
Else {
    $model = (Get-WmiObject Win32_computersystem -Property Model).value
    ConvertTo-Html -Body "<servertype>$model Physical Server</servertype>" >> "$filepath\$name.html"}

###Add Date of update
ConvertTo-Html -Body "<update>Last Updated: $date</update>" >> "$filepath\$name.html"



#Operating System information##
$osinfo = Get-WmiObject Win32_OperatingSystem -ComputerName $name  | Select Caption,CSDVersion,Version,OSArchitecture,Organization,InstallDate  
$osinfo.InstallDate = [management.managementDateTimeConverter]::ToDateTime($osinfo.InstallDate)
$osinfo | ConvertTo-html  -Body "<H2> Operating System Information </H2>" >> "$filepath\$name.html"

###Uptime and Last BootVolume
$wmiPerfOsSystem = Get-WmiObject -computer $name -class Win32_PerfFormattedData_PerfOS_System
$wmiOS = Get-WmiObject -computer $name -class Win32_OperatingSystem
$lastBoot = $wmiOS.ConvertToDateTime($wmiOS.LastBootUpTime)
New-TimeSpan -seconds $wmiPerfOsSystem.SystemUpTime -ErrorAction SilentlyContinue  |select days,hours,minutes,seconds | ConvertTo-html -body "<H2>Uptime</H2>" >> "$filepath\$name.html"
$lastBoot | select DateTime |ConvertTo-html -body "<H2>Last Boot Time</H2>" >> "$filepath\$name.html"


###Windows Updates last installed###
$Updates = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install" | select LastSuccessTime
$Updates.LastSuccessTime = [DATETIME]::Parse($Updates.LastSuccessTime)
$updatesout = New-Object -TypeName PSOBject
$Updates | ConvertTo-Html -body "<H2>Windows Updates Last Installed</H2>" >> "$filepath\$name.html"

###Event log
Get-EventLog -LogName System -Newest 30 -EntryType Error,warning | select TimeGenerated,entrytype,source,message | ConvertTo-html -body "<H2>System Events</H2>" >> "$filepath\$name.html"
Get-EventLog -LogName Application -Newest 30 -EntryType Error,warning | select TimeGenerated,entrytype,source,message | ConvertTo-html -body "<H2>Application Events</H2>" >> "$filepath\$name.html"


##Time Zone##
Get-WmiObject win32_TimeZone -ComputerName $name | Select caption  | ConvertTo-html  -Body "<H2> Time Zone </H2>" >> "$filepath\$name.html"

###Startup Applications###
Get-WmiObject win32_StartupCommand -ComputerName $name | select  caption,description,location,user | ConvertTo-html  -Body "<H2> Startup Applications </H2>" >> "$filepath\$name.html"

##local users###
$localusers = Get-WmiObject win32_UserAccount  -ComputerName $name -Filter "LocalAccount='$True'" | select Name,SID,PasswordExpires,Disabled,Lockout | ConvertTo-html  -Body "<H2> Local Users </H2>" >> "$filepath\$name.html"

###local user group membershipe
###Check if machine is a domain controller first to prevent large unreadable table
Import-Module servermanager
$Domaincontroller = get-windowsfeature  | where {$_.Name -eq "ADDS-Domain-Controller" -and $_.installed -eq $true}
if ($Domaincontroller -eq $null){
    $adsi = [ADSI]"WinNT://$env:COMPUTERNAME"
    $adsi.Children | where {$_.SchemaClassName -eq 'user'} | Foreach-Object {
    $groups = $_.Groups() | Foreach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
    $_ | Select-Object @{n='UserName';e={$_.Name}},@{n='Groups';e={$groups -join ';'}}
    } | ConvertTo-html  -Body "<H2> Local User Group Memberships </H2>" >> "$filepath\$name.html"
}
Else{}
# MotherBoard: Win32_BaseBoard # You can Also select Tag,Weight,Width 
Get-WmiObject -ComputerName $name  Win32_BaseBoard  |  Select Name,Manufacturer,Product,SerialNumber,Status  | ConvertTo-html  -Body "<H2> MotherBoard Information</H2>" >> "$filepath\$name.html"

# Battery uncomment if you want this information 
#Get-WmiObject Win32_Battery -ComputerName $name  | Select Caption,Name,DesignVoltage,DeviceID,EstimatedChargeRemaining,EstimatedRunTime  | ConvertTo-html  -Body "<H2> Battery Information</H2>" >> "$filepath\$name.html"

# BIOS
$bios = Get-WmiObject win32_bios -ComputerName $name  | Select Manufacturer,Name,Version,PrimaryBIOS,ReleaseDate,SMBIOSBIOSVersion,SMBIOSMajorVersion,SMBIOSMinorVersion  
$bios.Releasedate =[management.managementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
$bios | ConvertTo-html  -Body "<H2> BIOS Information </H2>" >> "$filepath\$name.html"

# CD ROM Drive
Get-WmiObject Win32_CDROMDrive -ComputerName $name  |  select Name,Drive,MediaLoaded,MediaType,MfrAssignedRevisionLevel  | ConvertTo-html  -Body "<H2> CD ROM Information</H2>" >> "$filepath\$name.html"

# System Info
Get-WmiObject Win32_ComputerSystemProduct -ComputerName $name  | Select Vendor,Version,Name,IdentifyingNumber,UUID  | ConvertTo-html  -Body "<H2> System Information </H2>" >> "$filepath\$name.html"

# Hard-Disk
Get-WmiObject win32_diskDrive -ComputerName $name  | select Model,SerialNumber,InterfaceType,Size,Partitions  | ConvertTo-html  -Body "<H2> Harddisk Information </H2>" >> "$filepath\$name.html"

#Volumes
$volumes = Get-WmiObject win32_volume -ComputerName $name  | where {$_.DriveType -eq 3} | select Name,Label,@{Name="Capacity (GB)";Expression={$_.Capacity}},@{Name="Blocksize (KB)";Expression={$_.Blocksize}},BootVolume
    foreach($volume in $volumes){
    $volume.'Blocksize (KB)' = [MATH]::Round($volume.'Blocksize (KB)' /1KB)
    $volume.'Capacity (GB)' = [MATH]::Round($volume.'Capacity (GB)' /1GB)
  }
$volumes | ConvertTo-html  -Body "<H2> Volume Information </H2>" >> "$filepath\$name.html"

###Shares###
Get-WmiObject win32_Share | select name,caption,path,status | ConvertTo-html  -Body "<H2> Shares </H2>" >> "$filepath\$name.html"


# NetWork Adapters -ComputerName $name
Get-WmiObject win32_networkadapter -ComputerName $name -Filter "Manufacturer != 'Microsoft'" | Select Name,Manufacturer,Description ,AdapterType,Speed,MACAddress,NetConnectionID |  ConvertTo-html  -Body "<H2> Network Card Information</H2>" >> "$filepath\$name.html"

###Network Adapter configuration
$netconfig = Get-WmiObject win32_networkadapterconfiguration -ComputerName $name -Filter "IPEnabled = True"| select IPAddress,DefaultIPGateway,IPSubnet,@{Name="PrimaryDNS";Expression ={$_.DNSServerSearchOrder}},DNSDomain,DHCPEnabled,Description 
foreach ($ip in $netconfig) {
    $ip.IPAddress = $ip.IPAddress[0]
    if ($ip.DefaultIPGateway -ne $null){
        $ip.DefaultIPGateway = $ip.DefaultIPGateway[0]
        }
    $ip.IPSubnet = $ip.IPSubnet[0]
    if ($ip.PrimaryDNS -ne $null){
        $ip | Add-Member -MemberType NoteProperty -Name "SecondaryDNS" -Value $ip.PrimaryDNS[1]
        $ip.PrimaryDNS = $ip.PrimaryDNS[0]}
    }
$netconfig | ConvertTo-Html -Body "<H2> Network Configuration </H2>" >> "$filepath\$name.html"

# Memory
$Memory = Get-WmiObject Win32_PhysicalMemory -ComputerName $name  | select BankLabel,DeviceLocator,@{Name="Capacity (GB)";Expression = {$_.Capacity}},Manufacturer,PartNumber,SerialNumber,Speed
foreach ($DIMM in $memory){
    $DIMM.'Capacity (GB)' = [MATH]::Round($DIMM.'Capacity (GB)' /1GB)
    }
$Memory | ConvertTo-html  -Body "<H2> Physical Memory Information</H2>" >> "$filepath\$name.html"

# Processor 
Get-WmiObject Win32_Processor -ComputerName $name  | Select Name,Manufacturer,Caption,DeviceID,MaxClockSpeed,CurrentVoltage,DataWidth,L2CacheSize,L3CacheSize,NumberOfCores,NumberOfLogicalProcessors,Status  | ConvertTo-html  -Body "<H2> CPU Information</H2>" >> "$filepath\$name.html"

## System enclosure 
Get-WmiObject Win32_SystemEnclosure -ComputerName $name  | Select Manufacturer,SerialNumber  | ConvertTo-html  -Body "<H2> System Enclosure Information </H2>" >> "$filepath\$name.html"

##Page File
$pagefiles = Get-WmiObject Win32_PageFileUsage -ComputerName $name  | Select Caption,@{Name="AllocatedBaseSize (GB)";Expression ={$_.AllocatedBaseSize}}
foreach ($pagefile in $pagefiles){
    $pagefile.'AllocatedBaseSize (GB)' = [MATH]::Round($pagefile.'AllocatedBaseSize (GB)' /1024)
}
$pagefiles| ConvertTo-html  -Body "<H2>Page File</H2>" >> "$filepath\$name.html"

####Windows Features
Import-Module servermanager
$features = get-windowsfeature | select DisplayName,Name,Installed,FeatureType 
foreach($feature in $features){
    $installed = $feature.Installed
    $installed = $installed.ToString()
    $installed = $installed.Replace("$false","")
    $installed = $installed.Replace("$true","X")
    $feature.installed = $installed
    }
$features | where{$_.installed -eq "X"} |ConvertTo-Html -Body " <H2>Installed Windows Features </H2>" >> "$filepath\$name.html"
$webserver = $features | where {$_.Name -eq "Web-Server"}

###IIS
if ($webserver.Installed -eq "X") {
    Import-Module webadministration

    $sites = Get-WebApplication
    $bindings = Get-WebBinding | select site,protocol,bindinginformation,certificatehash,certificatestorename,itemXpath
    $sites | select @{name="Name";expression={$_.path}},applicationpool,enabledprotocols,physicalpath,serviceautostartenabled  | ConvertTo-Html -Body " <H2>IIS Websites </H2>" >> "$filepath\$name.html"
    foreach ($binding in $bindings){
        #Clean up names to be suitable for display
        $applicationbind = $binding.ItemXPath
        $applicationbind = $applicationbind.Replace("/system.applicationHost/sites/site[@name='","")
        $applicationbind = $applicationbind.Replace("' and @id=","")
        $applicationbind = $applicationbind -replace "\'+[0123456789]+\'+\]",""
        $binding.site = $applicationbind
    }
    $bindings| select site,protocol,bindinginformation,certificatehash,certificatestorename | ConvertTo-Html -Body " <H2>IIS Bindings </H2>" >> "$filepath\$name.html"
    }
ELSE{}


## Installed Applications 
Get-WmiObject Win32_Product -ComputerName $name  | Select @{Name="Application Name";Expression={$_.Caption}},Version,Vendor  | sort 'Application Name' | ConvertTo-html  -Body "<H2>Installed Applications</H2>" >> "$filepath\$name.html"


##Services## 
Get-WmiObject win32_Service -ComputerName $name  | Select DisplayName,StartMode,State,PathName,StartName | Sort DisplayName  | ConvertTo-html  -Body "<H2>Services</H2>" >> "$filepath\$name.html"


###Create Local directory and copy needed Resources
if ((Test-Path $Scriptdir) -eq $false){
    mkdir $Scriptdir
    }
Else{}
Copy-Item "$filepath\Resources" -Recurse $Scriptdir -Force -Confirm:$false
## Invoke Expressons opens the report in browser
#invoke-Expression "$filepath\$name.html"

###Check if scheduled task option is enabled
if ($enableschedtask -eq $true){
    ####Make a copy og the script locaally
    Copy-Item "$MasterScriptVersion" "$Scriptdir" -Confirm:$false
    ####Check if scheduled task already installed
    $tasks = $null
    $tasks = schtasks /query /tn "Server Inventory"
    if ($tasks -eq $null){
        $password = $Scheduledtaskpass
        ###Create Scheduled Task to Run###
        Copy-Item "$MasterScriptVersion" "$Scriptdir" -Confirm:$false
        $text = "C:\windows\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy unrestricted -file $Scriptdir\ServerInventory.ps1"
        $TaskName = 'Server Inventory'
        $TaskRun = $text 
        #create task now
        schtasks.exe /create /sc 'daily' /ru $Scheduledtaskuser /rp $password /tn $Taskname /tr $TaskRun /st '04:00'
        }
        Else {write-host "skipping scheduled task already present"}
}
Else {write-host "skipping scheduled task setup"}
####Now Checking for ScheduledTasks
Set-Location $Scriptdir\Resources\
$schtasks = .\Get-ScheduledTask.ps1 -ComputerName $name | select TaskName,Enabled,ActionType,Action,LastRunTime,LastResult,State,Author,Created,RunAs 
$schtasks | ConvertTo-Html -Body "<H2>Scheduled Tasks</H2>" >> "$filepath\$name.html"

####Send Email if required
If ($sendemail -eq $true){
    Send-MailMessage -To $to -Subject $subject -From $from  $subject -SmtpServer $smtp -Priority "High" -BodyAsHtml -Attachments "$filepath\$name.html"}
