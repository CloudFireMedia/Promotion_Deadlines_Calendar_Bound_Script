var SCRIPT_NAME = 'Promotion Deadlines Calendar';
var SCRIPT_VERSION = 'v1.2.dev_ajr';

function onOpen(){ 

  var ui = SpreadsheetApp.getUi();

  ui.createMenu("[ CloudFire ]")
  
    .addItem("[ 1 ] Import Planning Calendar Data (this script is not yet written)", "") 
    .addItem("[ 2 ] Format Planning Calendar Data", "prepareNewYearsData") // TODO - doesn't do any formatting except to merge the WEEK column
    .addItem("[ 3 ] Apply Due Dates to New Data", "applyFormula") // is an arrayformula now - the action just fixes the formulas now
    .addItem("[ 4 ] Populate Bulletin Schedule Based on Previous Year's Schedule", "repopulateBulletins")
    
    .addSeparator()
    
/////////////////// GOT TO HERE ///////////////////    
    
    .addItem("Format Sheet", "PL.calendar_formatSheet")
    
    .addSubMenu(
      ui.createMenu('Format...')
      .addItem("Hide Old Rows", "PL.calendar_hideRows")
      .addItem("Delete Empty Rows", "PL.calendar_removeEmptyRows")
      .addItem("Set Border Color", "PL.calendar_colorBorders")
      .addItem("Set Weeks Format", "PL.calendar_setWeeksFormat")
      .addItem("Set Events Format", "PL.calendar_setEventsFormat")
    )
    
    .addSeparator()
    
    .addSubMenu(
      ui.createMenu('Tools')
      .addItem('Restore Header', 'PL.calendar_restoreHeader')
      .addItem('Backup Header', 'PL.calendar_backupHeader')
      .addSeparator()
      .addItem('Enable Automation', 'PL.calendar_setupAutomation') //note: do NOT run this from the library, use a proxy function ala: function setupAutomation(){PL.setupAutomation_responseForm()}
      .addItem('Disable Automation', 'PL.calendar_disableAutomation') //note: do NOT run this from the library, use a proxy function ala: function disableAutomation(){PL.disableAutomation_responseForm()}
    )  
    
    .addSeparator()
    
    .addItem("Check Deadlines", "PL.calendar_checkDeadlines")
    .addItem("Check for Errors", "PL.calendar_checkTeamSheetsForErrors")
    
    .addToUi();
    
} // onOpen()

function prepareNewYearsData() {PDC.prepareNewYearsData()}
function applyFormula() {PDC.applyFormula()}
function repopulateBulletins_() {PDC.repopulateBulletins_()}
/*
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
function () {PDC.()}
*/

function addRowBelow()               {PDC.addRowBelow()}
function deleteRow()                 {PDC.deleteRow()}
function calculateDeadlines()        {PDC.calculateDeadlines()}
function fillSponsor()               {PDC.fillSponsor()}
function fillTier()                  {PDC.fillTier()}
function processResponse(eventArray) {PDC.processResponse(eventArray)}
function newEventPopup()             {PDC.newEventPopup()}

function onEdit(e)                 {}
function onOpen(e)                 {PDC.onOpen()}
//function onEdit_Triggered(e)     {PL.onEdit_eventsCalendar_Triggered(e)}//not needed
function setupAutomation()         {PDC.setupAutomation_eventsCalendar()}
function disableAutomation()       {PDC.disableAutomation_eventsCalendar()}
function calendar_dailyTrigger()   {PDC.calendar_dailyTrigger()}