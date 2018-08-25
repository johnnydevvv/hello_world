var doc;
var profileList = [];
var spreadsheet;
var layoutSheet;
var layout;
var layoutvalue;
var profilevalue;
var objectlist = new Array();
var values = [];
var layouts = [];
var docObject;
var recordtype;

function createScripts (){
  
   getProfiles ();

  for (k in profileList){
    var profile = profileList[k];
    profilevalue = profile;
    
    createDoc();
  }
}
  
function createDoc (){
    
    doc = DocumentApp.create(profilevalue+" Profile Script");
    var body = doc.getBody();
    var header = body.appendParagraph("Test "+profilevalue+" Profile");
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  getObjects ();
  
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  var profileSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Profiles"));
  var lastRow = profileSheet.getLastRow();
  var range = profileSheet.getRange(1, 1, lastRow, 4);
  var values = range.getValues();
  
  for (i in objectlist) {
      var object = objectlist[i];
      docObject = object;
   
    var par1 = body.appendParagraph(object+"\r");
    par1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Navigate to "+object+" tab\r");
    
   for (j in values){
     layouts=[];
     var value = values[j][1];
     var profile = values[j][0];
   if(object==value&&profile==profilevalue) {
        layout=[values[j][2], values[j][3]];
        layouts.push(layout);
   }
        
  for (i in layouts) {
    recordtype =  layouts[i][0];
    layoutvalue = layouts[i][1];
                  
     //set body for section
     layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(docObject));
    if (recordtype=="") {body.appendParagraph("Click 'New'\r");
                         }
    else if (recordtype!=="") {body.appendParagraph("Click 'New' and select the record type below:\r\r"+recordtype+"\r");
                         }
     body.appendParagraph("Check that the fields in the following table are on the create new record form:\r");
     body.appendParagraph("R = Read Only, M = Mandatory, E = Editable\r");
     getCreateFields ();
     body.appendParagraph("\rSave the record and check that the following fields are visible in the detail view of the new record:\r");
     getVisibleFields ();
  }
   }
   } 
     }

function getProfiles (){
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  var profileSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Profiles'));
  var lastRow = profileSheet.getLastRow();
  var range = profileSheet.getRange(1, 1, lastRow, 1);
  var values = range.getValues();
  var profiles = new Array();
  for(i in values){
    var row = values[i];
    var duplicate = false;
    for(j in profiles){
      if(row.join() == profiles[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      profiles.push(row);
      profileList=profiles
    }
  }
  }
        
function getObjects (){
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  var profileSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Profiles'));
  var lastRow = profileSheet.getLastRow();
  var range = profileSheet.getRange(2, 2, lastRow, 1);
  var values = range.getValues();
  var objects = new Array();
  for(i in values){
    var row = values[i];
    var duplicate = false;
    for(j in objects){
      if(row.join() == objects[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      objects.push(row);
      objectlist=objects
    }
  }
  }

function getVisibleFields () {
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(docObject));
  var lastRow = layoutSheet.getLastRow();
  var lastColumn = layoutSheet.getLastColumn();
  var range = layoutSheet.getRange(1, 1, lastRow, lastColumn);
  var fieldValues = range.getValues();
  var columnvalue
  var fieldTable = [];
  for (i in fieldValues) {
      var column = fieldValues[0][i];
         if(column==layoutvalue){
         columnvalue = [i];
           for (j in fieldValues) {
             var tableHeader = fieldValues[0][columnvalue]
                var picklistColumn = (+columnvalue)+1
                picklistColumn = picklistColumn.toFixed(0);
                if (fieldValues[j][columnvalue]==tableHeader) {
                  fieldTable.push([fieldValues[j][0],fieldValues[j][columnvalue],fieldValues[j][picklistColumn],"Pass/Fail"]);
                }
             else if (fieldValues[j][columnvalue]!=="") {
                  fieldTable.push([fieldValues[j][0],fieldValues[j][columnvalue],fieldValues[j][picklistColumn],""]);
                }
  }
  }
}
  doc.getBody().appendTable(fieldTable);
}
function getCreateFields () {
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(docObject));
  var lastRow = layoutSheet.getLastRow();
  var lastColumn = layoutSheet.getLastColumn();
  var range = layoutSheet.getRange(1, 1, lastRow, lastColumn);
  var fieldValues = range.getValues();
  var columnvalue
  var fieldTable = [];
  for (i in fieldValues) {
      var column = fieldValues[0][i];
         if(column==layoutvalue){
         columnvalue = [i];
           for (j in fieldValues) {
             var tableHeader = fieldValues[0][columnvalue]
                var picklistColumn = (+columnvalue)+1
                picklistColumn = picklistColumn.toFixed(0);
                if (fieldValues[j][columnvalue]==tableHeader) {
                  fieldTable.push([fieldValues[j][0],fieldValues[j][columnvalue],fieldValues[j][picklistColumn],"Pass/Fail"]);
                }
             else if (fieldValues[j][columnvalue]=="M") {
                  fieldTable.push([fieldValues[j][0],fieldValues[j][columnvalue],fieldValues[j][picklistColumn],""]);
                }
             else if (fieldValues[j][columnvalue]=="E") {
                  fieldTable.push([fieldValues[j][0],fieldValues[j][columnvalue],fieldValues[j][picklistColumn],""]);
                }
  }
  }
}
  doc.getBody().appendTable(fieldTable);
}

