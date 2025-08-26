<# 
.SYNOPSIS
  Windows Machine Inventory as JSON (no HTML).
  Original idea/script by Alan Newingham – rewritten to emit JSON.

.DESCRIPTION
  Collects machine info via CIM/WMI and writes a single JSON file.

.PARAMETER OutputPath
  Folder to write the JSON file to (default $PSScriptRoot).

.PARAMETER IncludeInstalledProducts
  Includes Win32_Product (slow; may trigger MSI self-repair). Off by default.

.EXAMPLE
  .\PowershellSystemInventory-json.ps1
  .\PowershellSystemInventory-json.ps1 -OutputPath 'D:\inv' -IncludeInstalledProducts
#>

[CmdletBinding()]
param(
  [string]$OutputPath = $PSScriptRoot,
  [switch]$IncludeInstalledProducts
)

$ErrorActionPreference = 'Stop'

# Ensure output folder exists
if (-not (Test-Path -LiteralPath $OutputPath)) {
  New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$ComputerName = $env:COMPUTERNAME
$now = Get-Date

Write-Host "Collecting inventory on $ComputerName ..." -ForegroundColor Yellow

# --- General / System ---
$cs = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object `
  Model, Manufacturer, @{n='LocalAdministrator';e={$_.PrimaryOwnerName}}, @{n='SystemType';e={$_.SystemType}}

$boot = Get-CimInstance -ClassName Win32_BootConfiguration | Select-Object `
  Name, @{n='OSInstallLocation';e={$_.ConfigurationPath}}

$bios = Get-CimInstance -ClassName Win32_BIOS | Select-Object `
  Manufacturer, SerialNumber, @{n='BiosVersion';e={$_.SMBIOSBIOSVersion}}

$os = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object `
  SystemDirectory, @{n='Caption';e={$_.Caption}}, BuildNumber, Version, SerialNumber, InstallDate, LastBootUpTime, OSArchitecture

$tz = Get-CimInstance -ClassName Win32_TimeZone | Select-Object `
  Bias, @{n='Caption';e={$_.Caption}}, @{n='StandardName';e={$_.StandardName}}

# --- Storage ---
$logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' |
  Select-Object DeviceID,
                @{n='SizeGB';e={[int]($_.Size/1GB)}},
                @{n='FreeGB';e={[int]($_.FreeSpace/1GB)}}

$diskDrives = Get-CimInstance -ClassName Win32_DiskDrive |
  Where-Object MediaType -eq 'Fixed hard disk media' |
  Select-Object SystemName, @{n='Model';e={$_.Model}}, @{n='SizeGB';e={[int]($_.Size/1GB)}},
                InterfaceType, SerialNumber

# --- CPU / Memory ---
$cpu = Get-CimInstance -ClassName Win32_Processor | Select-Object `
  Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors, Status

$physicalMemory = Get-CimInstance -ClassName Win32_PhysicalMemory | ForEach-Object {
  [pscustomobject]@{
    PartNumber          = $_.PartNumber
    Tag                 = $_.Tag
    SerialNumber        = $_.SerialNumber
    Manufacturer        = $_.Manufacturer
    ConfiguredClockMHz  = $_.ConfiguredClockSpeed
    ConfiguredVoltagemV = $_.ConfiguredVoltage
    CapacityGB          = [math]::Round($_.Capacity/1GB, 1)
    BankLabel           = $_.BankLabel
    DeviceLocator       = $_.DeviceLocator
  }
}

# --- Network ---
$netAdaptersPS = @{}; Get-NetAdapter | ForEach-Object { $netAdaptersPS.Add($_.MacAddress, $_) }
$netAdapters = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' |
  Select-Object -Property Description, DHCPEnabled, DHCPServer,
                @{n='IPAddress';e={ $_.IPAddress -join ';' }},
                @{n='IPSubnet';e={ $_.IPSubnet -join ';' }},
                @{n='DefaultGateway';e={ $_.DefaultIPGateway -join ';' }},
                DNSDomain,
                @{n='DNSServerSearchOrder';e={ $_.DNSServerSearchOrder -join ';' }},
                MACAddress
				| ForEach-Object {
					[PSCustomObject] @{
					  Name = $netAdaptersPS[$_.MACAddress.Replace(':','-')].Name
					  Description = $_.Description
					  ifIndex = $netAdaptersPS[$_.MACAddress.Replace(':','-')].ifIndex
					  Status = $netAdaptersPS[$_.MACAddress.Replace(':','-')].Status
					  LinkSpeed = $netAdaptersPS[$_.MACAddress.Replace(':','-')].LinkSpeed
					  DHCPEnabled = $_.DHCPEnabled
					  DHCPServer = $_.DHCPServer
					  IPAddress = $_.IPAddress
					  IPSubnet = $_.IPSubnet
					  DefaultGateway = $_.DefaultGateway
					  DNSDomain = $_.DNSDomain
					  DNSServerSearchOrder = $_.DNSServerSearchOrder
					  MACAddress = $_.MACAddress
					}
				}

# --- Printers ---
$printers = Get-CimInstance -ClassName CIM_Printer | Select-Object `
  Name, DriverName, PrinterState, PrinterStatus, Location, PortName, Network, Shared, WorkOffline

# --- User profiles / hotfixes ---
$userProfiles = Get-ChildItem -Path 'C:\Users\' -ErrorAction SilentlyContinue |
  Select-Object Name, LastWriteTime, FullName

$hotfixes = Get-CimInstance -ClassName Win32_QuickFixEngineering |
  Select-Object HotFixID, Description, InstalledOn, InstalledBy

# --- Video / Monitors / USB ---
$video = Get-CimInstance Win32_VideoController | Select-Object `
  Status,
  @{n='Model';e={$_.Description}},
  @{n='AdapterRAM_GB';e={[math]::Round($_.AdapterRAM/1GB,1)}},
  DriverDate, DriverVersion, VideoModeDescription

$monitorCount = (Get-CimInstance Win32_VideoController).Count

$usbDevices = Get-PnpDevice -Class USB -PresentOnly -ErrorAction SilentlyContinue |
  Select-Object Class, Status, @{n='DeviceName';e={$_.FriendlyName}}, @{n='InstanceId';e={$_.InstanceId}}

# --- Last logged-on (by folder write) ---
$lastLog = Get-ChildItem 'C:\Users' -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1 -Property Name, LastWriteTime

# --- Optional (slow) Installed Products ---
$installedProducts = @()
if ($IncludeInstalledProducts) {
  try {
    $installedProducts = Get-CimInstance -ClassName Win32_Product |
      Select-Object Vendor, @{n='Name';e={$_.Name}}, Version, IdentifyingNumber, InstallDate
  } catch {
    Write-Warning "Win32_Product konnte nicht abgefragt werden: $($_.Exception.Message)"
  }
}

# --- Build final object ---
$inventory = [pscustomobject]@{
  Computer              = $ComputerName
  ReportVersion         = '0.9.0-json'
  GeneratedAt           = $now
  GeneratedBy           = $env:USERNAME

  General               = $cs
  BootConfiguration     = $boot
  BIOS                  = $bios
  OperatingSystem       = $os
  TimeZone              = $tz

  LogicalDisks          = $logicalDisks
  DiskDrives            = $diskDrives
  Processor             = $cpu
  PhysicalMemory        = $physicalMemory

  NetworkAdapters       = $netAdapters
  Printers              = $printers

  UserProfiles          = $userProfiles
  Hotfixes              = $hotfixes

  VideoControllers      = $video
  MonitorCount          = $monitorCount
  USBDevices            = $usbDevices

  LastUserFolderTouched = if ($lastLog) { $lastLog.Name } else { $null }

  InstalledProducts     = $installedProducts
}

# --- Write JSON ---
$jsonPath = Join-Path -Path $OutputPath -ChildPath "$ComputerName.json"
$inventory | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonPath -Encoding UTF8

Write-Host "JSON written to: $jsonPath" -ForegroundColor Green
