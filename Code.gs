var SCRIPT_NAME = 'Promotion Deadlines Calendar';
var SCRIPT_VERSION = 'v1.2.dev_ajr';

function onOpen(){ 

  var ui = SpreadsheetApp.getUi();

  // SpreadsheetApp.getUi().createMenu(caption).addItem(caption, functionName).addSeparator().addToUi()

  ui.createMenu("CloudFire")
    .addItem("[ 1 ] Import Planning Calendar Data (this script is not yet written)", "importCalendarData")
    .addItem("[ 2 ] Format Planning Calendar Data", "prepareNewYearsData") // TODO - doesn't do any formatting except to merge the WEEK column
    .addItem("[ 3 ] Apply Due Dates to New Data", "applyFormula") // is an arrayformula now - the action just fixes the formulas now
    .addItem("[ 4 ] Populate Bulletin Schedule Based on Previous Year's Schedule", "repopulateBulletins")
    .addSeparator()
    .addItem("Format Sheet", "formatSheet")
    .addSubMenu(
      ui.createMenu('Format...')
      .addItem("Hide Old Rows", "hideRows")
      .addItem("Delete Empty Rows", "removeEmptyRows")
      .addItem("Set Border Color", "colorBorders")
      .addItem("Set Weeks Format", "setWeeksFormat")
      .addItem("Set Events Format", "setEventsFormat")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Tools')
      .addItem('Restore Header', 'restoreHeader')
      .addItem('Backup Header', 'backupHeader')
      .addSeparator()
      .addItem('Enable Automation', 'setupAutomation') 
      .addItem('Disable Automation', 'disableAutomation') 
    ) 
    .addSeparator()
    .addItem("Check Deadlines", "checkDeadlines")
    .addItem("Check for Errors", "checkTeamSheetsForErrors")
    .addSeparator()
    .addItem("Add row below", "addRowBelow")
    .addItem("Delete row", "deleteRow")
    .addItem("Calculate deadlines", "calculateDeadlines")
    .addItem("New event popup", "newEventPopup")
    .addToUi();
    
} // onOpen()

// Menu items
function importCalendarData()        {} // TODO
function prepareNewYearsData()       {return PDC. prepareNewYearsData()}
function applyFormula()              {return PDC. applyFormula()}
function repopulateBulletins()       {return PDC. repopulateBulletins()}
function formatSheet()               {return PDC. formatSheet()}
function hideRows()                  {return PDC. hideRows()}
function removeEmptyRows()           {return PDC. removeEmptyRows()}
function colorBorders()              {return PDC. colorBorders()}
function setEventsFormat()           {return PDC. setEventsFormat()}
function restoreHeader()             {return PDC. restoreHeader()}
function backupHeader()              {return PDC. backupHeader()}
function setupAutomation()           {return PDC. setupAutomation()}
function disableAutomation()         {return PDC. disableAutomation()}
function checkDeadlines()            {return PDC. checkDeadlines()}
function checkTeamSheetsForErrors()  {return PDC. checkTeamSheetsForErrors()}

// Menu items and macros
function addRowBelow()               {return PDC. addRowBelow()}
function deleteRow()                 {return PDC. deleteRow()}
function calculateDeadlines()        {return PDC. calculateDeadlines()}
function newEventPopup()             {return PDC. newEventPopup()}

// Triggers
function dailyTrigger()              {return PDC.dailyTrigger()}

// Client Side 
function fillSponsor()               {return PDC.fillSponsor()}
function fillTier()                  {return PDC.fillTier()}