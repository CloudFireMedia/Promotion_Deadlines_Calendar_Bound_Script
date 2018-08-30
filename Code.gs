var SCRIPT_NAME = ' CCN Events Promotion Calendar';
var SCRIPT_VERSION = 'v1.2.dev_ajr'

//settings
var documentConfig = {
  dataSheetName : 'Communications Director Master',
};

//push config options to library
PDC.updateConfig({eventsCalendar:documentConfig});

function onEdit(e)                 {}
function runMeToGrantPermissions() {}
function onOpen(e)                 {PDC.onOpen()}
//function onEdit_Triggered(e)     {PL.onEdit_eventsCalendar_Triggered(e)}//not needed
function setupAutomation()         {PDC.setupAutomation_eventsCalendar()}
function disableAutomation()       {PDC.disableAutomation_eventsCalendar()}
function calendar_dailyTrigger()   {PDC.calendar_dailyTrigger()}

//debug
function dumpConfig() {PDC.dumpConfig()}
function foo() {PDC.foo()}

// Original bound code now in ChCS Github Repo - https://github.com/Christ-Church-Nashville/ChCS_Events_Promotion_Calendar v1.0
