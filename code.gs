/*
Author: Moksea
Description: This script manages cooldown timers for various in-game activities in a spreadsheet and sends notifications to Discord when the timers reach 0.
*/

// Function to get the Discord webhook URL from cell K2
function getWebhookURL() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var webhookURL = sheet.getRange("K3").getValue();
  return webhookURL;
}

// Function to send a notification to Discord
function sendNotification(message) {
  const webhookURL = getWebhookURL();
  const payload = {
    content: message
  };
  
  try {
    // Send a POST request to the webhook URL
    UrlFetchApp.fetch(webhookURL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });
  } catch (error) {
    console.error("Error sending notification:", error);
  }
}

// onEdit function triggered when a cell is edited
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (range.getColumn() == 6 && range.getValue() === true) { 
    var eventRow = range.getRow();
    var cooldownHours = sheet.getRange(eventRow, 3).getValue();
    if (cooldownHours && !isNaN(cooldownHours)) {
      var now = new Date();
      var resetTime = new Date(now.getTime() + cooldownHours * 60 * 60 * 1000); // Set the reset time based on cooldown hours
      var formattedDateTime = Utilities.formatDate(resetTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(eventRow, 4).setValue(formattedDateTime);
      sheet.getRange(eventRow, 5).setValue(""); // Clear the countdown timer cell to force update
      Logger.log("Reset triggered for row " + eventRow + ", new reset time: " + formattedDateTime);
      updateCountdowns();
    }
    range.setValue(false); // Reset the checkbox back to FALSE
  }
}

// Function to update countdown timers
function updateCountdowns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] != "") {
      try {
        var activityName = data[i][1]; // Assuming activity name is in the second column
        var resetTime = new Date(data[i][3]);
        var diff = resetTime.getTime() - now.getTime();
        if (diff > 0) {
          var hours = Math.floor(diff / (60 * 60 * 1000));
          var minutes = Math.floor((diff % (60 * 60 * 1000)) / (60 * 1000));
          var seconds = Math.floor((diff % (60 * 1000)) / 1000);
          var countdown = hours + "h " + minutes + "m " + seconds + "s";
          sheet.getRange(i + 1, 5).setValue(countdown);
        } else {
          sheet.getRange(i + 1, 5).setValue("Ready");
          sheet.getRange(i + 1, 4).setValue(""); // Set Completion Time to blank
          sendNotification("Cooldown timer reached 0 for " + activityName);
        }
      } catch (e) {
        console.error("Error processing row " + (i + 1) + ": " + e.message);
      }
    }
  }
}

// Function to create time trigger
function createTimeTrigger() {
  ScriptApp.newTrigger('updateCountdowns')
    .timeBased()
    .everyMinutes(1)
    .create();
}
