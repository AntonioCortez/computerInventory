'06/2016 
'Antonio Cortez
'Base script.
'Retrieving important computer Information
'References
'https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-computersystem
'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem



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

  WScript.Echo _ 
  "Name: " & objItem.Name & vbCrLf & _
  "Manufacturer: " & objItem.Manufacturer & vbCrLf & _
  "Model: " & objItem.Model & vbCrLf & _
  "User name: " & objItem.UserName & vbCrLf & _
  "System type: " & objItem.PCSystemType  & vbCrLf & _
  "System type2: " & objItem.SystemType & vbCrLf & _
  "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory & vbCrLf 
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
	"SO TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize & vbCrLf & _
	"SO name: " & objItem.Name 
Next
