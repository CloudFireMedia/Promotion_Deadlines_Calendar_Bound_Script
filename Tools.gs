function backupHeader_(){
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  if( ! SpreadsheetApp.getActiveSheet().getName() == sheet.getName()){
    if(Browser.msgBox('Backup Header', "\
Only the "+sheet.getName()+' sheet can be backed up.\\n\\n\
Switch to it?\
', Browser.Buttons.OK_CANCEL)
      != 'ok') return;
    sheet.activate();
    SpreadsheetApp.flush();
  }
  
  //have to have frozen rows to know what to backup
  if( ! sheet.getFrozenRows()){
    Browser.msgBox('Backup Header', "There does not appear to be a header defined.\\nPlease use View -> Freeze to set the rows you want as the header then run this again.", Browser.Buttons.OK)
    return;
  }
  
 //prompt to overwrite existing backup
  if(Browser.msgBox('Backup Header', "\
This will backup the header rows so they can be restored later\\n\
mostly in case the formulas are accidentally deleted.\\n\
The frozen rows at the top are considered the header \\n\
so be certain all the header rows are frozen..  \\n\\n\
Continue to backup header (and replace existing backup if any)?\
", Browser.Buttons.OK_CANCEL) != 'ok') return;
  
  var range = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getLastColumn());
  var header = {
    values : range.getValues(),
    formulae : range.getFormulas(),
    backgrounds : range.getBackgrounds(),
    fontColors : range.getFontColors(),
    fontFamilies : range.getFontFamilies(),
    fontSizes : range.getFontSizes(),
    fontStyles : range.getFontStyles(),
    fontWeights : range.getFontWeights(),
    fontLines : range.getFontLines(),
    horizontalAlignments : range.getHorizontalAlignments(),
//    mergedRanges : range.getMergedRanges(),//nope, no this way anyway.  Doesn't serialize.  Need to write something to handle this.
    notes : range.getNotes(),
    numberFormats : range.getNumberFormats(),
    verticalAlignments : range.getVerticalAlignments(),
    wraps : range.getWraps(),
  }
  PropertiesService.getScriptProperties().setProperty('epc:backup_header', JSON.stringify(header))
  Browser.msgBox('Backup Header', "Backup complete.", Browser.Buttons.OK)
}

function restoreHeader_() {

  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);
  
  //make sure we're on the right sheet
  if( ! SpreadsheetApp.getActiveSheet().getName() == sheet.getName()){
    if(Browser.msgBox('Restore Header', "\
Only the "+sheet.getName()+' sheet can be restored.\\n\\n\
Switch to it?\
', Browser.Buttons.OK_CANCEL)
      != 'ok') return;
    sheet.activate();
    SpreadsheetApp.flush();
  }

  //check to see if there is a backup
  var backup = JSON.parse(PropertiesService.getScriptProperties().getProperty('epc:backup_header'));
  if( ! backup){
    Browser.msgBox('Restore Header', "There is no header backup to restore.", Browser.Buttons.OK);
    return;
  }

  var headerRows = backup.values.length;

  //verify overrite current header
  if(Browser.msgBox('Restore Header', "\
This will restore the header from the last backup\\n\
overwriting the first "+headerRows+" rows in the sheet in the process. \\n\
If that would replace data, stop and add additional rows for the header.  \\n\\n\
Continue to retore header and replace existing rows?\
", Browser.Buttons.YES_NO) != 'yes') return;

  //proceed with restoring header
  var range = sheet.getRange(1, 1, backup.values.length, backup.values[0].length);
  //merge formulae into values as setFormulas() overwrites values even when they are null
  backup.values = backup.values.map(function(c, i){ return c.map(function(c2, i2){
    //replace value with formula when exits
    return backup.formulae[i][i2] ? backup.formulae[i][i2] : c2 || null;
  }) });
  range.setValues(backup.values);//set combined values and formulae
  range.setNotes(backup.notes);
  range.setWraps(backup.wraps);
  range.setBackgrounds(backup.backgrounds);
  range.setNumberFormats(backup.numberFormats);
  range.setVerticalAlignments(backup.verticalAlignments);
  range.setHorizontalAlignments(backup.horizontalAlignments);
  range.setFontSizes(backup.fontSizes);
  range.setFontLines(backup.fontLines);
  range.setFontColors(backup.fontColors);
  range.setFontStyles(backup.fontStyles);
  range.setFontWeights(backup.fontWeights);
  range.setFontFamilies(backup.fontFamilies);
  
  sheet.setFrozenRows(headerRows);
  
  Browser.msgBox('Restore Header', "Header restored.", Browser.Buttons.OK);
}
