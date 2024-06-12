// Initialization and Setup Functions

// Initialize varaibles
let currentDateTime = new Date(); // Store current currentDateTime
//let timeZone = Session.getScriptTimeZone(); // Store user's timezone
//Logger.log(timeZone); // Log user's timezone
//let localTime = Utilities.formatDate(time, timeZone, 'yyyy-MM-dd HH:mm:ss'); // format currentDateTime
let sheet = getActiveSheet(); // store current sheet

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Haiku')
      .addItem('Setup Trigger', 'setupTriggers')
      .addItem('Update', 'update')
      .addItem('Test Webhook', 'testWebhook')
      .addToUi();
     Logger.log("Menu created successfully");
  } catch (error) {
    Logger.log("Error in onOpen: " + error.message);
  }
}

// Utility Functions
function getActiveSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function update() {
  try {
    updateLocalTime();
    updateServerTime();
    updateCountdowns(sheet);
  } catch (error) {
    Logger.log("Error in update: " + error.message);
  }
}

// Main Functions
function onEdit(e) {
  try {
    Logger.log("onEdit triggered");
    
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    Logger.log("Edited range: " + range.getA1Notation());

    if (range.getColumn() == 6 && range.getValue() === true) { 
      Logger.log("Checkbox in column 6 checked");

      var eventRow = range.getRow();
      var cooldownHours = sheet.getRange(eventRow, 3).getValue();
      Logger.log("Cooldown hours: " + cooldownHours);

      if (cooldownHours && !isNaN(cooldownHours)) {

        var resetTime = new Date(currentDateTime);
        resetTime.setHours(resetTime.getHours() + 1 + cooldownHours);
        Logger.log(resetTime);

        sheet.getRange(eventRow, 4).setValue(resetTime);
        sheet.getRange(eventRow, 5).setValue(""); 
        Logger.log("Reset triggered for row " + eventRow + ", new reset time: " + resetTime);
        updateCountdowns(sheet);
      }
      range.setValue(false); 
    }
  } catch (error) {
    Logger.log("Error in onEdit: " + error.message);
  }
}

function updateCountdowns(sheet) {
  try {
    Logger.log("updateCountdowns triggered");

    var data = sheet.getDataRange().getValues();
//  var now = new Date();
//  var userTimeZone = getUserTimeZone(sheet);
    var batchUpdates = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][3] != "" && data[i][3] != "Target Time") {
        try {
          var activityName = data[i][1]; 
          var resetTime = new Date(data[i][3]);
          var diff = resetTime.getTime() - currentDateTime.getTime();
          var countdown;

          if (diff > 0) {
            var hours = Math.floor(diff / (60 * 60 * 1000));
            var minutes = Math.floor((diff % (60 * 60 * 1000)) / (60 * 1000));
            var seconds = Math.floor((diff % (60 * 1000)) / 1000);
            countdown = hours + "h " + minutes + "m " + seconds + "s";
          } else {
            countdown = "Ready";
            sheet.getRange(i + 1, 4).setValue(""); 
            sendNotification("Cooldown timer reached for " + activityName);
          }

          batchUpdates.push({range: sheet.getRange(i + 1, 5), value: countdown});
        } catch (e) {
          Logger.log("Error processing row " + (i + 1) + ": " + e.message);
        }
      }
    }

    batchUpdates.forEach(update => update.range.setValue(update.value));
    Logger.log("Countdowns updated successfully");
  } catch (error) {
    Logger.log("Error in updateCountdowns: " + error.message);
  }
}

function updateLocalTime() {
  try {
    sheet.getRange("localTime").setValue(currentDateTime);
    Logger.log("Local time displayed successfully");
  } catch (error) {
    Logger.log("Error in showLocalTime: " + error.message);
  }
}

function updateServerTime() {
  try {
    var serverTimeZone = 'GMT+2';
    serverTime = Utilities.formatDate(currentDateTime, serverTimeZone, 'HH:mm:ss')
    sheet.getRange("serverTime").setValue(serverTime);
    Logger.log("Server time updated successfully");
  } catch (error) {
    Logger.log("Error in updateServerTime: " + error.message);
  }
}