// Initialization and Setup Functions

// Initialize varaibles
let time = new Date(); // Store current time
let timeZone = Session.getScriptTimeZone(); // Store user's timezone
Logger.log(timeZone); // Log user's timezone
let localTime = Utilities.formatDate(time, timeZone, 'yyyy-MM-dd HH:mm:ss'); // format datetime
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
    updateCountdowns(sheet);
    updateServerTime(sheet);
    updateLocalTime(sheet);
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

        var resetTime = new Date(localTime);
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
          var diff = resetTime.getTime() - now.getTime();
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

function updateServerTime(sheet) {
  try {
    var serverTimeCell = sheet.getRange("serverTime");
    var serverTime = new Date();
    var serverTimeOffset = serverTime.getTimezoneOffset(); 
    var gmtPlus2Time = new Date(serverTime.getTime() + (serverTimeOffset + 120) * 60 * 1000); 
    serverTimeCell.setValue(Utilities.formatDate(gmtPlus2Time, Session.getScriptTimeZone(), "HH:mm:ss"));
    Logger.log("Server time updated successfully");
  } catch (error) {
    Logger.log("Error in updateServerTime: " + error.message);
  }
}

function sendNotification(message) {
  try {
//  var sheet = getActiveSheet();
    const webhookURL = getWebhookURL(sheet);
    const payload = { content: message };
    UrlFetchApp.fetch(webhookURL, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
    Logger.log("Notification sent: " + message);
  } catch (error) {
    Logger.log("Error sending notification: " + error.message);
  }
}

// Trigger Management
function setupTriggers() {
  try {
    resetTriggers();
    Logger.log("Triggers set up successfully");
  } catch (error) {
    Logger.log("Error in setupTriggers: " + error.message);
  }
}

function resetTriggers() {
  try {
    deleteAllTriggers();
    createTimeTrigger();
    Logger.log("Triggers reset successfully");
  } catch (error) {
    Logger.log("Error in resetTriggers: " + error.message);
  }
}

function getExistingTriggers(functionName) {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var existingTriggers = triggers.filter(trigger => trigger.getHandlerFunction() === functionName);
    Logger.log("Existing triggers for " + functionName + ": " + existingTriggers.length);
    return existingTriggers;
  } catch (error) {
    Logger.log("Error in getExistingTriggers: " + error.message);
  }
}

function deleteAllTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    Logger.log("All triggers deleted successfully");
  } catch (error) {
    Logger.log("Error deleting triggers: " + error.message);
  }
}

function createTimeTrigger() {
  try {
    var existingTriggers = getExistingTriggers('update');
    if (existingTriggers.length === 0) {
      ScriptApp.newTrigger('update').timeBased().everyMinutes(1).create();
      Logger.log("Created time trigger for 'update'");
    } else {
      Logger.log("Time trigger for 'update' already exists");
    }
  } catch (error) {
    Logger.log("Error creating time trigger: " + error.message);
  }
}