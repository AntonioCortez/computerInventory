'06/2016 
'Antonio Cortez @kacorius
'Base script.
'Retrieving important computer Information
'References
'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-physicalmemory
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-computersystem
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisk

Wscript.Echo _ 
 "Date: " & FormatDateTime(now,2) & vbCrLf & _
 "Time: " & FormatDateTime(now,4) & vbCrLf


Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'List BIOS information
Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
For each objBIOS in colBIOS
    Wscript.Echo "Build Number: " & objBIOS.BuildNumber & vbCrLf & _
	"Current Language: " & objBIOS.CurrentLanguage & vbCrLf & _
	"Installable Languages: " & objBIOS.InstallableLanguages & vbCrLf & _
	"Manufacturer: " & objBIOS.Manufacturer & vbCrLf & _
	"BIOS Name & version: " & objBIOS.Name & vbCrLf & _
	"Primary BIOS: " & objBIOS.PrimaryBIOS & vbCrLf & _
	"Release Date: " & objBIOS.ReleaseDate & vbCrLf & _
	"Serial Number: " & objBIOS.SerialNumber & vbCrLf & _
	"SMBIOS Version: " & objBIOS.SMBIOSBIOSVersion & vbCrLf & _
	"SMBIOS Major Version: " & objBIOS.SMBIOSMajorVersion & vbCrLf & _
	"SMBIOS Minor Version: " & objBIOS.SMBIOSMinorVersion & vbCrLf & _
	"SMBIOS Present: " & objBIOS.SMBIOSPresent & vbCrLf & _
	"Status: " & objBIOS.Status & vbCrLf & _
	"Version: " & objBIOS.Version & vbCrLf 
Next

'List Fisical Memory data
Dim Total,factor,typem
Total = 0
Set colRAM = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
For each objRAM in colRAM

	Select case objRAM.FormFactor
		case 0 factor="Unknown"
		case 1 factor="Other"
		case 2 factor="SIP"
		case 3 factor="DIP"
		case 4 factor="ZIP"
		case 5 factor="SOJ"
		case 6 factor="Proprietary"
		case 7 factor="SIMM"
		case 8 factor="DIMM"
		case 9 factor="TSOP"
		case 10 factor="PGA"
		case 11 factor="RIMM"
		case 12 factor="SO-DIMM"
		case 13 factor="SRIMM"
		case 14 factor="SMD"
		case 15 factor="SSMP"
		case 16 factor="QFP"
		case 17 factor="TQFP"
		case 18 factor="SOIC"
		case 19 factor="LCC"
		case 20 factor="PLCC"
		case 21 factor="BGA"
		case 22 factor="FPBGA"
		case 23 factor="LGA"
	End Select

	Select case objRAM.MemoryType
		case 0 typem="Unknown"
		case 1 typem="Other"
		case 2 typem="DRAM "
		case 3 typem="Synchronous DRAM"
		case 4 typem="CACHE RAM"
		case 5 typem="EDO"
		case 6 typem="EDRAM"
		case 7 typem="VRAM"
		case 8 typem="SRAM"
		case 9 typem="RAM"
		case 10 typem="ROM"
		case 11 typem="FLASH"
		case 12 typem="EEPROM"
		case 13 typem="FEPROM"
		case 14 typem="EPROM"
		case 15 typem="CDRAM"
		case 16 typem="3DRAM"
		case 17 typem="SDRAM"
		case 18 typem="SGRAM"
		case 19 typem="RDRAM"
		case 20 typem="DDR"
		case 21 typem="DDR2"
		case 22 typem="DDR2 FB-DIMM"
		case 24 typem="DDR3"
		case 25 typem="FBD2"
		case 26 typem="DDR4"
	End Select


	Wscript.Echo vbCrLf & _
	"RAM DeviceLocator: " & objRAM.DeviceLocator & vbCrLf & _
	"RAM Tag: " & objRAM.Tag & vbCrLf & _
	"RAM FormFactor: " & factor & vbCrLf & _
	"RAM MemoryType: " & typem & vbCrLf & _
	"RAM Model: " & objRAM.Model & vbCrLf & _
	"RAM Capacity: "   & Round(objRAM.Capacity/1073741824,2) & vbCrLf & _
	"RAM Speed: " & objRAM.Speed & vbCrLf & _
	"RAM Manufacturer: " & objRAM.Manufacturer & vbCrLf & _
	"RAM Description: " & objRAM.Description
	Total=Total+Round(objRAM.Capacity/1073741824,2)
Next
WScript.Echo vbCrLf & "Total physical installed memory: " & Total & vbCrLf

'List Drive C information 
Set colPStorage = objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DeviceID = 'C:'")
For Each objPStorage in colPStorage
  WScript.Echo _ 
  "C Size: " & Round(objPStorage.Size/1073741824,2) & vbCrLf & _
  "C VolumeName: " & objPStorage.VolumeName & vbCrLf & _
  "C FreeSpace: " & Round(CDbl(objPStorage.FreeSpace)/1073741824,2) & vbCrLf 
Next

' List processor information
Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor in colProcessors
  WScript.Echo _ 
  "Manufacturer: " & objProcessor.Manufacturer & vbCrLf & _
  "Name: " & objProcessor.Name & vbCrLf & _
  "Description: " & objProcessor.Description & vbCrLf & _
  "Processor ID: " & objProcessor.ProcessorID & vbCrLf & _
  "Address Width: " & objProcessor.AddressWidth & vbCrLf & _
  "Data Width: " & objProcessor.DataWidth & vbCrLf & _
  "Family: " & objProcessor.Family & vbCrLf & _
  "Maximum Clock Speed: " & objProcessor.MaxClockSpeed & vbCrLf 
Next

' List Computer Manufacturer and Model information
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
	
	Select case objItem.PCSystemType
		case 0 typem="Unspecified"
		case 1 typem="Desktop"
		case 2 typem="Mobile"
		case 3 typem="Workstation"
		case 4 typem="Enterprise Server"
		case 5 typem="Small Office and Home Office (SOHO) Server"
		case 6 typem="Appliance PC"
		case 7 typem="Performance Server"
		case 8 typem="Maximum"
	End Select

	WScript.Echo _ 
	"Name: " & objItem.Name & vbCrLf & _
	"Manufacturer: " & objItem.Manufacturer & vbCrLf & _
	"Model: " & objItem.Model & vbCrLf & _
	"User name: " & objItem.UserName & vbCrLf & _
	"System type: " & typem  & vbCrLf & _
	"System type2: " & objItem.SystemType & vbCrLf & _
	"GB TotalPhysicalMemory available after OS: " & objItem.TotalPhysicalMemory/1073741824 & vbCrLf 
Next

'List SO information
Set solItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem In solItems
	WScript.Echo _ 
	"Buld number: " & objItem.BuildNumber & vbCrLf & _
	"SO SerialNumber: " & objItem.SerialNumber & vbCrLf & _
	"SO Caption: " & objItem.Caption & vbCrLf & _
	"SO version: " & objItem.Version & vbCrLf & _
	"SO CSDVersion: " & objItem.CSDVersion & vbCrLf & _
	"GB SO TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize/1048576 & vbCrLf & _
	"SO name: " & objItem.Name 
Next
