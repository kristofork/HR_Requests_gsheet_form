function onFormSubmit(e) {
  // Log spreadsheet and sheet objects for debugging
  Logger.log("Spreadsheet object (ss):", e.source);
  Logger.log("Sheet object (sheet):", e.source.getSheetByName("HR Request Form Responses"));

  var sheet = e.source.getSheetByName("HR Request Form Responses");  // Replace with your actual sheet name
  if (!sheet) {
    Logger.log("Error: Sheet not found!");
    return;  // Exit if sheet is not found
  }

  // Wait for 1 second to allow potential data processing (optional)
  Utilities.sleep(1000);  // Adjust sleep time as needed

  var lastRow = sheet.getLastRow();

  var typeOfRequest = sheet.getRange(lastRow, 4).getValue();  // Get "Type of Request" value

  var targetSheetName, targetRange;

  switch (typeOfRequest) {
    case "Contractor Hiring Request":
      targetSheetName = "Hiring";
      targetRange = "A:M";
      break;
    case "Partner Account Request":
      targetSheetName = "Access";
      targetRange = "A:D,N:X";
      break;
    case "Term Notification":
      targetSheetName = "Term";
      targetRange = "A:D,Y:AI";
      break;
    default:
      return;  // Exit if request type doesn't match criteria
  }

 var targetSheet = e.source.getSheetByName(targetSheetName);
  var targetRow = targetSheet.getLastRow() + 1;

  var sourceRange = sheet.getRange(lastRow, 1, 1, targetRange.split(":")[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1);  // Adjust for dynamic range
  var targetRange = targetSheet.getRange(targetRow, 1, 1, sourceRange.getNumColumns());

  targetRange.setValues(sourceRange.getValues());
}