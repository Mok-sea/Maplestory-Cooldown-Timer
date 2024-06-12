// Initialization and Setup Functions

// Initialize varaibles
let time = new Date(); // Store current time
let timeZone = Session.getScriptTimeZone(); // Store user's timezone
Logger.log(timeZone); // Log user's timezone
let localTime = Utilities.formatDate(time, timeZone, 'yyyy-MM-dd HH:mm:ss'); // format datetime
let sheet = getActiveSheet(); // store current sheet


// Utility Functions
function getActiveSheet() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function updateLocalTime() {
    try {
//  var sheet = getActiveSheet();
//  var userTimeZone = getUserTimeZone(sheet);
  
//  var now = new Date();
//  var localTime = Utilities.formatDate(now, userTimeZone, "HH:mm:ss");
  
    sheet.getRange("localTime").setValue(localTime);
      Logger.log("Local time displayed successfully");
    } catch (error) {
      Logger.log("Error in showLocalTime: " + error.message);
    }
}
  