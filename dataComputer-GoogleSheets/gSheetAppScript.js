/*
Antonio Cortez
References:
https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
https://developers.google.com/apps-script/reference/spreadsheet/sheet?hl=en
https://developers.google.com/apps-script/guides/html?hl=en
https://developers.google.com/apps-script/guides/web?hl=en

To see the result on the GoogleSheets go to:
https://docs.google.com/spreadsheets/d/1ipubH2umjV9uGwv-ws3YBUd7-xAFQuLLGt90ciuH0Fc/edit?usp=sharing

*/
function doGet(e) {
    //Displays the text on the webpage.
    //return "chocho";
    return ContentService.createTextOutput("This is a GET Request!");

}

function doPost(e) {
  
 var data = JSON.parse(e.postData.contents);
 var sheet = SpreadsheetApp.getActiveSheet();

 sheet.appendRow([
    data.Date,
    data.Time,
    data.Hardware.SN,
    data.Hardware.Manufacturer,
    data.Hardware.Model,
    data.Hardware.PCtype,
    data.Hardware.BPC,
    data.Hardware.Processor,
    data.Hardware.RAM,
    data.Hardware.Storage.Tot,
    data.Hardware.Storage.Free,
    data.Software.SO,
    data.Software.ver,
    data.Cname
  ]);
  
  //return ContentService.createTextOutput("This is a POST Request 28h! + data.time: "+ data.Time +" + data.Cname: "+data.Cname+ "   + e.postData.contents: "+e.postData.contents);
  return ContentService.createTextOutput("100A");
}