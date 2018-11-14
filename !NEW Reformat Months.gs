// resets month rows formatting in Col E by reseting the the month array in Col A, 
// this is necessary due to a quirk in the way that conditional formatting works on 
// merged rows
function onEdit(e) {
   var commsDirectorMasterSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();   

/*    NOTE: 
      var commsDirectorMasterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
      would literally specifies 'Communications Director Master' sheet as var CommsDirectorMasterSheet, 
      but the method is slightly slower than the getActiveSheet method; and Comms Director Master 
      sheet should not need to be specified, as all other sheets are read-only. 
      --cbd 2018.10.28
*/

   var editRange = { // D4:G184
    top : 4,
    bottom : commsDirectorMasterSheet.getMaxRows(),
    left : 5,
    right : 7
  };
  var thisRow = e.range.getRow();
  if (thisRow < editRange.top || thisRow > editRange.bottom) return;
  var thisCol = e.range.getColumn();
  if (thisCol < editRange.left || thisCol > editRange.right) return;
//  var commsDirectorMasterSheet = e.range.getSheet();
  commsDirectorMasterSheet.getRange("A4:A4")   
    .clearContent(); // Clear cell of values (not formatting)
}
