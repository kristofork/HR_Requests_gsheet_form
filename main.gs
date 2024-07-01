function hrFormSubmit(e) {

// Set up variables
var targetSheetName, targetRange, targetOffset, slackWebHookURL;


// Targetoffset is needed because the columns in the source sheet are different than the target sheet. The offset is to place them correctly in the current sheet.
// Adding additional fields to the source sheet will cause data to not be aligned correctly in the target sheets.
targetOffset = 0;

// set data that was submitted via the form to variable
var rawData = e;

// set the request type to variable
var typeOfRequest = rawData['namedValues']['Type of Request'][0];

// get the active sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// get last row of sheet
var lastRow = sheet.getLastRow();

// validate sheet exists
  if (!sheet) {
    Logger.log("Error: Sheet not found!");
    return;  // Exit if sheet is not found
  }

  // Wait for 1 second to allow potential data processing (optional)
  Utilities.sleep(1000);  // Adjust sleep time as needed

  switch (typeOfRequest) {
    case "Contractor Hiring Request":
      targetSheetName = "Hiring";
      targetRange = ['A:M'];
      slackWebHookURL = "https://URL_FROM_SLACK_WEB_HOOK";
      break;
    case "Partner Account Request":
      targetSheetName = "Access";
      targetRange = ['A:D', 'N:X'];
      targetOffset = 9
      slackWebHookURL = "https://URL_FROM_SLACK_WEB_HOOK"
      break;
    case "Term Notification":
      targetSheetName = "Term";
      targetRange = ['A:D','Y:AI'];
      targetOffset = 20
      slackWebHookURL = "https://URL_FROM_SLACK_WEB_HOOK"
      break;
    default:
      return;  // Exit if request type doesn't match criteria
  }

  // Get target sheet by ID
  var tsheet = SpreadsheetApp.openById("1CNA9lfvyZNBlGRQIl9WjqMm18JTh7et4Pr9QbPyCdpc")
  var targetSheet = tsheet.getSheetByName(targetSheetName)
  var targetRow = targetSheet.getLastRow() + 1;

  // Loop through the columns array to get the range of the last row
  for (var i = 0; i < targetRange.length; i++) {
    // Get the column range
    var rangeStr = targetRange[i];
    //var range = sheet.getRange(targetRange[i]);
    var rangeCols = rangeStr.split(":");

      // Get the starting and ending column indices
    var startCol = columnLetterToNumber(rangeCols[0]);
    var endCol = columnLetterToNumber(rangeCols[1]);
    var startColTarget = startCol- targetOffset;
    var endColTarget = endCol - targetOffset;

    // Get the range of the last row for the current column range
    var range = sheet.getRange(lastRow, startCol, 1, endCol - startCol + 1);
    var values = range.getValues();
    console.log(values);

    // Set values in the target sheet
    if (i >= 1)
    {
      var targetSheetRange = targetSheet.getRange(targetRow, startColTarget, 1, endColTarget - startColTarget + 1);
    }
    else
    {
      var targetSheetRange = targetSheet.getRange(targetRow, startCol, 1, endCol - startCol + 1);
    }
    targetSheetRange.setValues(values);
  }

// Web hook to Slack
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var form = FormApp.openByUrl(spreadsheet.getFormUrl());
var allResponses = form.getResponses();
var latestResponse = allResponses[allResponses.length - 1];
var response = latestResponse.getItemResponses();
var payload = {};
    for (var i = 0; i < response.length; i++) {
        var question = response[i].getItem().getTitle();
        var cleanquestion = question.replace(/[\\\/\s\W_]/g, '');
        var answer = response[i].getResponse();
        payload[cleanquestion] = answer;
    }
  
    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload)
    };
     Logger.log('Sending webhook to Slack: ' + JSON.stringify(payload));
UrlFetchApp.fetch(slackWebHookURL, options);

}

// Helper function to convert column letters to column numbers
function columnLetterToNumber(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return column;
}
