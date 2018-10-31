var SCRIPT_NAME = 'Promotion Deadlines Calendar';
var SCRIPT_VERSION = 'v1.2.dev_ajr';

function onOpen(){ 

  var ui = SpreadsheetApp.getUi();

  ui.createMenu("CloudFire")
    .addItem("Format Sheet", "formatSheet")
    .addSeparator()
    .addItem('Restore Header', 'restoreHeader')
    .addItem('Backup Header', 'backupHeader')
    .addSeparator()
    .addItem('Enable Automation', 'setupAutomation') 
    .addItem('Disable Automation', 'disableAutomation') 
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

// Menu items - Formatting
function formatSheet()               {return PDC.formatSheet()}

// Menu items - Tools
function restoreHeader()             {return PDC.restoreHeader()}
function backupHeader()              {return PDC.backupHeader()}

function setupAutomation()           {return PDC.setupAutomation()}
function disableAutomation()         {return PDC.disableAutomation()}

function checkDeadlines()            {return PDC.checkDeadlines()}
function checkTeamSheetsForErrors()  {return PDC.checkTeamSheetsForErrors()}

// Menu items and macros
function addRowBelow()               {return PDC.addRowBelow()}
function deleteRow()                 {return PDC.deleteRow()}
function calculateDeadlines()        {return PDC.calculateDeadlines()}
function newEventPopup()             {return PDC.newEventPopup()}

// Triggers
function dailyTrigger()              {return PDC.dailyTrigger()}

// Client Side 
function fillSponsor()               {return PDC.fillSponsor()}
function fillTier()                  {return PDC.fillTier()}