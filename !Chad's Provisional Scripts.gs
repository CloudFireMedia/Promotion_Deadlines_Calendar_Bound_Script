//macro menu item 'Add Row Below' (shortcut COMMAND + OPTION + SHIFT + 2)
function AddRowBelow() {
  var flag = 0;
   SpreadsheetApp.getActiveSpreadsheet().toast("Working...","",-1);
  var spreadsheet = SpreadsheetApp.getActive();
  var target= spreadsheet.getActiveRange().getValues();
  var srow = spreadsheet.getCurrentCell().getRow();
  
  if(target.length>1){
    Browser.msgBox('Error! \\n\\nYou have selected more than one row. Please select only \\none cell or row and run the function again.');
  }
  else{
  if((target[0]+"").indexOf(',')==-1){
  
  spreadsheet.getRange(srow+':'+srow).activate();
  }
  var flag = 1;
   spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getCurrentCell().offset(0, 3).activate();
  spreadsheet.getCurrentCell().offset(-1, 0).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 2).activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('No')
  .setTextStyle(0, 2, SpreadsheetApp.newTextStyle()
  .setFontSize(7)
  .build())
  .build());
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('No')
  .setTextStyle(0, 2, SpreadsheetApp.newTextStyle()
  .setFontSize(7)
  .build())
  .build());
  spreadsheet.getCurrentCell().offset(0, 2).activate();
  spreadsheet.getCurrentCell().offset(-1, 0, 1, 3).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 3).activate();
  spreadsheet.getCurrentCell().offset(-1, 0).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, -8).activate();
  while(spreadsheet.getRange('A'+(srow)).getValue()=='')
  {
    srow =srow-1;
  }
  
  var erow = spreadsheet.getCurrentCell().getRow();
  var val2 = spreadsheet.getRange('A'+(erow+1)).getValue();
  if(val2 != '') {
   spreadsheet.getRange('A' + srow + ':'+ 'A'+ erow ).activate().merge();
   spreadsheet.getRange('B' + srow + ':'+ 'B'+ erow ).activate().mergeVertically();
  }     
    spreadsheet.getRange('D'+erow).activate();
  SpreadsheetApp.getActiveSpreadsheet().toast("Row added.");
}

};



//macro menu item 'Delete Row' (shortcut COMMAND + OPTION + SHIFT + 1)
function DeleteRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  SpreadsheetApp.getActiveSpreadsheet().toast("Row deleted.");
};



//menu item 'Initialize Promotion Deadlines'
function calculateDeadlines(){
  var spreadsheet = SpreadsheetApp.getActive();
   SpreadsheetApp.getActiveSpreadsheet().toast("Working...","",-1);
   // Prompt for the value of bronze
  var b =  SpreadsheetApp.getUi().prompt("Please enter deadline for Bronze Promotion in number of weeks").getResponseText();
  var s =  SpreadsheetApp.getUi().prompt("Please enter deadline for Silver Promotion in number of weeks.").getResponseText();
  var g =  SpreadsheetApp.getUi().prompt("Please enter deadline for Gold Promotion in number of weeks.").getResponseText();
  PropertiesService.getScriptProperties().setProperty('b', b);
  PropertiesService.getScriptProperties().setProperty('s', s);
  PropertiesService.getScriptProperties().setProperty('g', g);
  var lastRow = spreadsheet.getLastRow();
  Browser.msgBox('Total number of rows are ' + lastRow);
  var target_range = spreadsheet.getRange('D4:D'+lastRow);
 var  values = target_range.getValues(); 
 for (var c = 4; c < lastRow+1; c++) {

  var d  = spreadsheet.getRange('D' + c ).getValue();

   //finding value for column I
  var result = new Date(d.getTime()-7*b*3600000*24);
   
  var day = (result+"").substring(0,3);
   //Browser.msgBox(day);
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
   
  }
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('I' + c ).setValue(result);
   
   //Finding Value for column J
     
      var result = new Date(d.getTime()-7*s*3600000*24);
 
  var day = (result+"").substring(0,3);
   //Browser.msgBox(day);
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
  }
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('J' + c ).setValue(result);

   //finding value for Column K
   
    var result = new Date(d.getTime()-7*g*3600000*24);
 
  var day = (result+"").substring(0,3);
   //Browser.msgBox(day);
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
  }
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('K' + c ).setValue(result);
      }
  SpreadsheetApp.getActiveSpreadsheet().toast("Success! Bronze, Silver, and Gold Promotion Deadlines have been assigned to all events in the current sheet.");

};

function onEdit(e) 
{
  var spreadsheet = SpreadsheetApp.getActive();
  var thisCol = e.range.getColumn();
  
  if (thisCol != 4) return;
 // SpreadsheetApp.getActiveSpreadsheet().toast("Working...","",-1);
  var b=   PropertiesService.getScriptProperties().getProperty('b');
  var s =  PropertiesService.getScriptProperties().getProperty('s');
  var g = PropertiesService.getScriptProperties().getProperty('g');
  if(b == null){
    Browser.msgBox('Error! \\n\\nDefault values for Bronze, Silver and Gold Promotion Deadlines have not been assigned, and deadlines for the selected row cannot be automatically updated. \\n\\nPlease run \'Macros > Initialize Promotion Deadlines\' to assign default Promotion Deadlines values.');
  return;
  }
 
  var row = spreadsheet.getCurrentCell().getRow();

  var d  = spreadsheet.getRange('D' + row ).getValue();

   //finding value for column I
  var result = new Date(d.getTime()-7*b*3600000*24);
   
  var day = (result+"").substring(0,3);
   //if day is sunday move it forward to monday
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
   
  }
  //if day is saturday move it backward to friday
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('I' + row ).setValue(result);
   
   //Finding Value for column J
     
      var result = new Date(d.getTime()-7*s*3600000*24);
 
  var day = (result+"").substring(0,3);
   //Browser.msgBox(day);
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
  }
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('J' + row ).setValue(result);

   //finding value for Column K
   
    var result = new Date(d.getTime()-7*g*3600000*24);
 
  var day = (result+"").substring(0,3);
   //Browser.msgBox(day);
  if(day=='Sun'){
    var result = new Date(result.getTime()+1*3600000*24);
  }
  if(day=='Sat'){
    var result = new Date(result.getTime()-1*3600000*24);
  }

   spreadsheet.getRange('K' + row ).setValue(result);
   SpreadsheetApp.getActiveSpreadsheet().toast("Success! Promotion Deadlines have been updated for the selected row.");
  
} 