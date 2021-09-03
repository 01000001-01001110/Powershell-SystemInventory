<#
	.SYNOPSIS
		Windows Machine Inventory Using PowerShell.
    Written by - Alan Newingham
    Date: 9/3/2021

	.DESCRIPTION
		This script is to document the Windows machine. This script will work only for Local Machine.

	.EXAMPLE
		./inventory.ps1

	.OUTPUTS
		HTML File Output ReportDate , General Information , BIOS Information etc.

#>
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#Set-ExecutionPolicy RemoteSigned -ErrorAction SilentlyContinue

$ComputerName = (Get-Item env:\Computername).Value
$filepath = "c:\temp"



#CSS injection
$head = @"

<style>

body { 
  background-color:#9ca3ae;
  font-family:Verdana;
  font-size:12pt; 
  font-color: Black;
}

th { 
  font-family:Verdana;
  color:white;
  background-color:#094f9a;
  }

td, th { 
  font-family:Verdana;
  border:1px solid #056ecd; 
  border-collapse:collapse;
  white-space:pre; 
}

table, tr, td, th { 
  font-family:Verdana;
  padding: 2px;
  margin: 0px ;
  white-space:pre; 
}

tr:nth-child(odd) {
  font-family:Verdana;
  background-color: #cdd1d6;
}

tr:nth-child(even) {
  font-family:Verdana;
  background-color: #ebedef;
}

table {
  font-family:Verdana;
  width:95%;
  margin:0 auto 20px;
}

h1 {
  font-family:Verdana;
  color: #ffffff;
  text-align: center;
  width:95%;
}

h2 {
font-family:Verdana;
color:#6D7B8D;
}

caption {
  font-family:Verdana;
  text-align: left
}

.bannah {
  color: #071934;
  font-family:Verdana;
  font-size:14pt;
}

.footer 
{ 
background-color:#51cf66; 
margin-top:40px;
font-family: verdana;
font-size:10pt;
font-style:italic;
width:95%;
border:1px solid #056ecd;
}
</style>
"@



Write-Host "Executing Inventory Report!!! Please Wait !!!" -ForegroundColor Yellow 

#General Information
$ComputerSystem = Get-CimInstance -Class Win32_ComputerSystem | Select-Object -Property Model , Manufacturer , 
@{Name='Local Administrator';Expression={$_.PrimaryOwnerName -join '; '}},
@{Name='64 / 32 Bit';Expression={$_.SystemType -join '; '}}  |ConvertTo-Html -Fragment
$ComputerSystem = $ComputerSystem -replace "<table>", "<table>
<caption>General Information</caption>"
$ComputerSystem = $ComputerSystem -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Boot Configuration
$BootConfiguration = Get-CimInstance -Class Win32_BootConfiguration | Select-Object -Property Name , 
@{Name='OS Install Location';Expression={$_.ConfigurationPath -join '; '}}  | ConvertTo-Html -Fragment 
$BootConfiguration = $BootConfiguration -replace "<table>", "<table>
<caption>Boot Configuration</caption>"
$BootConfiguration = $BootConfiguration -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#BIOS Information
$BIOS = Get-CimInstance -Class Win32_BIOS | Select-Object -Property Manufacturer, SerialNumber, @{Name='BIOS Version';Expression={$_.SMBIOSBIOSVersion -join '; '}} | ConvertTo-Html -Fragment
$BIOS = $BIOS -replace "<table>", "<table>
<caption>BIOS Information</caption>"
$BIOS = $BIOS -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Operating System Information
$OS = Get-CimInstance -Class Win32_OperatingSystem | Select-Object -Property SystemDirectory, 
@{Name='Operating System';Expression={$_.Caption -join '; '}}, BuildNumber, Version, SerialNumber | ConvertTo-Html -Fragment
$OS = $OS -replace "<table>", "<table>
<caption>Operating System Information</caption>"
$OS = $OS -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Time Zone Information
$TimeZone = Get-CimInstance -Class Win32_TimeZone | Select-Object Bias,
@{Name='Time Zone';Expression={$_.Caption -join '; '}}, 
@{Name='Standard Name';Expression={$_.StandardName -join '; '}} | ConvertTo-Html -Fragment
$TimeZone = $TimeZone -replace "<table>", "<table>
<caption>Time Zone Information</caption>"
$TimeZone = $TimeZone -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Logical Disk Information
$Disk = Get-CimInstance -Class Win32_LogicalDisk -Filter DriveType=3 | Select-Object DeviceID , @{Expression={$_.Size /1Gb -as [int]};Label="Total Size(GB)"},@{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Size (GB)"} | ConvertTo-Html -Fragment
$Disk = $Disk -replace "<table>", "<table>
<caption>Hard Disk Capacity</caption>"
$Disk = $Disk -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Get drive information and display it. Make/Model.
$Drive = Get-CimInstance win32_diskdrive | Where-Object MediaType -eq 'Fixed hard disk media' | Select-Object SystemName,
@{Name='Hard Disk Model';Expression={$_.Model -join '; '}}, @{Name='Size(GB)';Exp={$_.Size /1gb -as [int]}} | ConvertTo-Html -Fragment
$Drive = $Drive -replace "<table>", "<table>
<caption>Hard Disk Information</caption>"
$Drive = $Drive -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#CPU Information
$SystemProcessor = Get-CimInstance -Class Win32_Processor  | Select-Object Name , 
@{Name='Processor Clock Speed';Expression={$_.MaxClockSpeed -join '; '}} , Manufacturer , 
@{Name='Processor Status';Expression={$_.status -join '; '}} | ConvertTo-Html -Fragment
$SystemProcessor = $SystemProcessor -replace "<table>", "<table>
<caption>Processor Information</caption>"
$SystemProcessor = $SystemProcessor -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Memory Information
$PhysicalMemory = Get-CimInstance -Class Win32_PhysicalMemory | Select-Object -Property PartNumber, 
@{Name='Memory Slot Occupied';Expression={$_.Tag -join '; '}}, SerialNumber  , Manufacturer , ConfiguredClockSpeed , ConfiguredVoltage , @{Name="Capacity(GB)";Expression={"{0:N1}" -f ($_.Capacity/1GB)}} | ConvertTo-Html -Fragment
$PhysicalMemory = $PhysicalMemory -replace "<table>", "<table>
<caption>Installed RAM Information</caption>"
$PhysicalMemory = $PhysicalMemory -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Network Information
$Network = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter DHCPEnabled=TRUE  | 
    Select-Object Description, DHCPServer, 
        @{Name='Ip Address';Expression={$_.IpAddress -join '; '}}, 
        @{Name='Ip Subnet';Expression={$_.IpSubnet -join '; '}}, 
        @{Name='Default Gateway';Expression={$_.DefaultIPgateway -join '; '}}, 
        @{Name='DNS Domain Name';Expression={$_.DNSDomain -join '; '}}, 
        @{Name='Primary DNS';Expression={$_.WinsPrimaryServer -join '; '}}, 
        @{Name='Secondary DNS';Expression={$_.WINSSecondaryServer -join '; '}} | ConvertTo-Html -Fragment 
        $Network = $Network -replace "<table>", "<table>
        <caption>Enabled Network Adapter(s) Information</caption>"
        $Network = $Network -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Printer Information
#Get-CimInstance -Class Win32_PrinterConfiguration | ConvertTo-Html -Fragment
#$Printer = Get-Printer -ComputerName $ComputerName | Select-Object -Property Name , DriverName | ConvertTo-Html -Fragment
$Printer = Get-CimInstance -ClassName CIM_Printer | Select-Object -Property Name , 
@{Name='Printer Driver Associated';Expression={$_.DriverName -join '; '}},
@{Name='Printer State';Expression={$_.PrinterState -join '; '}}, 
@{Name='Printer Status';Expression={$_.PrinterStatus -join '; '}}, 
@{Name='Remote Connection Method';Expression={$_.Location -join '; '}} | ConvertTo-Html -Fragment
$Printer = $Printer -replace "<table>", "<table>
<caption>Installed Printer Information</caption>"
$Printer = $Printer -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#User Accounts
$Directories = Get-ChildItem -Path "C:\Users\" | Select-Object -Property Name , 
@{Name='Last Directory Write';Expression={$_.LastWriteTime -join '; '}} | ConvertTo-Html -Fragment
$Directories = $Directories -replace "<table>", "<table>
<caption>User Profiles Loaded</caption>"
$Directories = $Directories -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Installed Windows Updates
$HotFix = Get-CimInstance -Class Win32_QuickFixEngineering | Select-Object -Property HotFixID |ConvertTo-Html -Fragment 
$HotFix = $HotFix -replace "<table>", "<table>
<caption>Hotfix Information (Mainly for Servers)</caption>"
$HotFix = $HotFix -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Installed Products
$InstalledProduct = Get-CimInstance -Class Win32_Product | Select-Object -Property Vendor, 
@{Name='Installed Software Title';Expression={$_.Name -join '; '}}, Version, 
@{Name='Registry Key Association';Expression={$_.IdentifyingNumber -join '; '}} | ConvertTo-HTML -Fragment
$InstalledProduct = $InstalledProduct -replace "<table>", "<table>
<caption>Installed Software</caption>"
$InstalledProduct = $InstalledProduct -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

#Monitors Installed
$monitors = Get-CimInstance Win32_VideoController 
$monitors=$monitors.count

#ReportDate
#$ReportDate = Get-Date | Select-Object -Property DateTime |ConvertTo-Html -Fragment

$videocard = Get-CimInstance Win32_VideoController | Select-Object -Property Status,
@{Name='Video Chipset Model';Expression={$_.Description -join '; '}}, 
@{Name='Memory(GB)';Expression={"{0:N1}" -f ($_.AdapterRAM/1GB)}}, 
@{Name='Driver Install Date';Expression={$_.DriverDate -join '; '}}, 
@{Name='Driver Version';Expression={$_.DriverVersion -join '; '}}, 
@{Name='Screensize';Expression={$_.VideoModeDescription -join '; '}}  | ConvertTo-HTML -Fragment
$videocard = $videocard -replace "<table>", "<table>
<caption>Video Card Information</caption>"
$videocard = $videocard -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

$USBDevices = Get-PnpDevice -Class USB -PresentOnly | Select-Object -Property Class, Status, 
@{Name='Device Name';Expression={$_.FriendlyName -join '; '}}, 
@{Name='Hardware Address';Expression={$_.InstanceId -join '; '}} | ConvertTo-HTML -Fragment
$USBDevices = $USBDevices -replace "<table>", "<table>
<caption>Active USB Devices</caption>"
$USBDevices = $USBDevices -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

$lastloggedon = Get-ChildItem "C:\Users" | Select-Object -Property Mode, Name, LastWriteTime | Sort-Object LastWriteTime -Descending | Select-Object  Mode, 
@{Name='Last Logged On User';Expression={$_.Name -join '; '}}, 
@{Name='Date / Time';Expression={$_.LastWriteTime -join '; '}} -first 1 | ConvertTo-HTML -Fragment
$lastloggedon = $lastloggedon -replace "<table>", "<table>
<caption>Last User Login</caption>"
$lastloggedon = $lastloggedon -replace "<colgroup><col/><col/><col/><col/></colgroup>", ""

$postContent = "<center><div class=""footer"">
<p>This report was generated On; $(get-date) by $((Get-Item env:\username).Value) on computer <b>$((Get-Item env:\Computername).Value)</b> <BR>Report Version 0.8.4, Written By: Alan Newingham</p>
</div></center>"

ConvertTo-Html -Head $head -Body "<center><p><H1><center>Computer: $ComputerName Information Export</center></H1></p></center>

$ComputerSystem 
$SystemProcessor
$PhysicalMemory
$BootConfiguration
$BIOS 
$OS
$Drive
$Disk
$Directories
$HotFix
$InstalledProduct
$Printer
$Network
<table> <colgroup><col/><col/></colgroup> <tr><th>Monitor Count</th></tr> <tr><td><center>$Monitors</center></td></tr></table>
$videocard
$USBDevices
$TimeZone
$lastloggedon
$postContent" -Title "$ComputerName"  | Out-File "$FilePath\$ComputerName.html" 

#Uncomment Invoke-Item line to have html popup when completed with inventory.
#Invoke-Item -Path "$FilePath\$ComputerName.html"
Write-Host "Script Execution Completed" -ForegroundColor Yellow
