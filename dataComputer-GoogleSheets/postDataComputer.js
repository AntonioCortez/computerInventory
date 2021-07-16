/*
06/2016
Antonio Cortez @kacorius
Jscript version
Send computer information to Google Sheets
https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page
https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios
https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor
https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-computersystem
https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem

XMLHTTPREQUEST:
https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms755436(v=vs.85)
*/

var data = {
    "Date" : "", 
    "Time" : "",
     "Cname": "",
    "Hardware" : {
        "SN": "",
        "Manufacturer" : "",
        "Model" : "",
        "PCtype" : "",
        "BPC" : "",
        "Processor" : "",
        "RAM" : "",
        "Storage" : {
            "Tot" : "",
            "Free" : ""
        }
    },
    "Software" : {
        "SO" : "",
        "ver" : ""
    }
} ;

var d = new Date();
data.Date = d.getDate()+"/"+(d.getMonth()+1)+"/"+d.getFullYear();
data.Time = d.getHours()+":"+d.getMinutes();

var objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2");

//BIOS information
var items= objWMIService.ExecQuery("Select * from Win32_BIOS");
data.Hardware.SN= items.ItemIndex(0).SerialNumber;

//List Fisical Memory data
items= objWMIService.ExecQuery("Select * from Win32_PhysicalMemory");
var oItems= new Enumerator(items); var ram=0;
for(;!oItems.atEnd();oItems.moveNext()){
    ram= ram + Math.round(oItems.item().Capacity/1073741824);
}
data.Hardware.RAM= ram + " GB"

//List Drive C information
items= objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DeviceID = 'C:'");
data.Hardware.Storage.Tot= Math.floor(items.ItemIndex(0).Size/1073741824) + " GB";
data.Hardware.Storage.Free= Math.floor(items.ItemIndex(0).FreeSpace/1073741824) + " GB";


//List processor information
items= objWMIService.ExecQuery("Select * from Win32_Processor");
data.Hardware.Processor= items.ItemIndex(0).Name;


//List Computer Manufacturer and Model information
var typem="";
items= objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem");
data.Cname= items.ItemIndex(0).UserName;
data.Hardware.Manufacturer= items.ItemIndex(0).Manufacturer;
data.Hardware.Model= items.ItemIndex(0).Model;

switch(items.ItemIndex(0).PCSystemType){
    case 0 : typem="Unspecified";  break;
    case 1 : typem="Desktop";  break;
    case 2 : typem="Mobile";  break;
    case 3 : typem="Workstation";  break;
    case 4 : typem="Enterprise Server";  break;
    case 5 : typem="Small Office and Home Office (SOHO) Server";  break;
    case 6 : typem="Appliance PC";  break;
    case 7 : typem="Performance Server";  break;
    case 8 : typem="Maximum";  break;
}

data.Hardware.PCtype= typem;
data.Hardware.BPC= items.ItemIndex(0).SystemType;

//List SO information
items= objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem");
data.Software.SO= items.ItemIndex(0).Caption;
data.Software.ver= items.ItemIndex(0).Version;

//making UpData
var upData='{"Date" : "'+data.Date+'", "Time" : "'+data.Time+'", "Cname": "'+data.Cname.replace('\\','\\\\')+'", "Hardware" : { "SN": "'+data.Hardware.SN+'", "Manufacturer" : "'+data.Hardware.Manufacturer+'", "Model" : "'+data.Hardware.Model+'", "PCtype" : "'+data.Hardware.PCtype+'", "BPC" : "'+data.Hardware.BPC+'", "Processor" : "'+data.Hardware.Processor+'","RAM" : "'+data.Hardware.RAM+'", "Storage" : { "Tot" : "'+data.Hardware.Storage.Tot+'", "Free" : "'+data.Hardware.Storage.Free+'"} }, "Software" : { "SO" : "'+data.Software.SO+'", "ver" : "'+data.Software.ver+'" } }';

WScript.Echo("brrr: "+upData+'\n');

//Upload data
var url= "https://script.google.com/macros/s/AKfycbzCbsFGhcVIPM0lqGUDjbHbXycmOsLQvhBXl3JD8lML6hZfnm4UW-zTw_wKdzCNHaeP/exec";

var http = new ActiveXObject("Msxml2.ServerXMLHTTP.6.0");
http.open("POST", url, false);
http.setRequestHeader("Content-type", "application/json; charset=utf-8");
http.send(upData);
var respon= http.responseText;
WScript.Echo(" "+respon);
