var SCRIPT_NAME = 'Promotion Deadlines Calendar';
var SCRIPT_VERSION = 'v1.5';

// Macros
function addRowBelow()        {return PDC.addRowBelow()}
function deleteRow()          {return PDC.deleteRow()}
function calculateDeadlines() {return PDC.calculateDeadlines()}
function newEventPopup()      {return PDC.newEventPopup()}

// Client-side callbacks
function fillSponsor(args)     {return PDC.fillSponsor(args)}
function fillTier(args)        {return PDC.fillTier(args)}
function processResponse(args) {return PDC.processResponse(args)}