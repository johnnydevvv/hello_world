var doc;
var spreadsheet;
var layoutSheet;
var layout;
var layoutvalue;
var object = "Account"
var profilevalue = "ENZ Standard User"

function createDoc() {
  doc = DocumentApp.create('Sample Document');
    
  Logger.clear();
  
  var app = SpreadsheetApp;
  spreadsheet = app.getActiveSpreadsheet();
  var profileSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet1'));
  var values = [];
  
  var body = doc.getBody();
  var header = body.appendParagraph("Test ENZ Standard User Profile");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Navigate to")
  body.appendParagraph("");
  body.appendParagraph(object);
  
  //get record types for profile
  for(var i=2; i<10; i++){
  var profile = profileSheet.getRange(i, 1).getValue();
  var objectmatch = profileSheet.getRange(i, 2).getValue();
  if(profile==profilevalue&&objectmatch==object){
     var recordtypes = profileSheet.getRange(i, 3).getValue();
    values.push(recordtypes);
    
  }
  }
  
   values.forEach(function(recordtype){
     for(var i=2; i<10; i++){
       var profile = profileSheet.getRange(i, 1).getValue();
       var recordtypecolumn = profileSheet.getRange(i, 3).getValue();
       if(profile==profilevalue&&recordtypecolumn==recordtype){
         layoutvalue = profileSheet.getRange(i, 4).getValue();
     Logger.log(layout);
     //set body for section
     layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet2'));
     var body = doc.getBody();
     body.appendParagraph("");
     body.appendParagraph("Click 'New' and select the record type below:");
     body.appendParagraph("");
     body.appendParagraph(recordtype);
     body.appendParagraph("");
     body.appendParagraph("Check that the following fields are on the create new record form:");
     body.appendParagraph("R = Read Only, M = Mandatory, E = Editable");
     body.appendParagraph("");
     getcreatefieldsforlayoutindex();
     body.appendParagraph("");
     body.appendParagraph("Save the record and check that the following fields are visible in the detail view of the new record:");
     body.appendParagraph("");
     getvisiblefieldsforlayoutindex();
     body.appendParagraph("");
  }
     }  
   }     
)
     }
     
function getvisiblefieldsforlayoutindex() {
  layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet2'));
  var columnvalue
  for(var i=1; i<100; i++){
  var columns = layoutSheet.getRange(1, i).getValue();
  if(columns==layoutvalue){
     columnvalue = layoutSheet.getRange(1, i).getColumn();
     columnvalue = columnvalue.toFixed(0);
Logger.log(columnvalue);
}
     }
     for(var i=2; i<8; i++){
  var visible = layoutSheet.getRange(i, columnvalue).getValue();
  var visiblefields = [];
       Logger.log(visible);
       var doesValueMatch = false; 
       if(visible==='M') {doesValueMatch=true;}
       else if(visible==='E'){doesValueMatch=true}
       else if(visible==='RO'){ doesValueMatch=true} 
       if(doesValueMatch) {
     var field = layoutSheet.getRange(i, 1).getValue();
   visiblefields.push(field);
    doc.getBody().appendParagraph(visiblefields);
  
  }
     }

     
}
function getcreatefieldsforlayoutindex() {
  layoutSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet2'));
  var columnvalue
  for(var i=1; i<100; i++){
  var columns = layoutSheet.getRange(1, i).getValue();
  if(columns==layoutvalue){
     columnvalue = layoutSheet.getRange(1, i).getColumn();
     columnvalue = columnvalue.toFixed(0);

}
     }
     for(var i=2; i<8; i++){
  var create = layoutSheet.getRange(i, columnvalue).getValue();
  var createfields = [];
       Logger.log(create);
       var doesValueMatch = false; 
       if(create==='M') {doesValueMatch=true;}
       else if(create==='E'){doesValueMatch=true}
       else if(create==='RO'){ doesValueMatch=true} 
       if(doesValueMatch) {
     var field = layoutSheet.getRange(i, 1).getValue();
     var tablevalue = field+" - "+create
   createfields.push(tablevalue);
   doc.getBody().appendParagraph(createfields);

  }
     }

     
}
