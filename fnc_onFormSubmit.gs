function onFormSubmit(e) {
  let rangeToCopy = e.range;
  let columns = rangeToCopy.getNumColumns();
  let rowToCopy = rangeToCopy.getRow();
  
  let worksheet = SpreadsheetApp.getActiveSpreadsheet();
  let responseSheet = worksheet.getSheetByName("Reimbursement Form Responses");
  
  let responses = e.values;
  let respondentName = responses[1];

  if (!responseSheet) {
    Logger.log("Main sheet 'Reimbursement Form Responses' not found");
    return;
  }
  
  
  // Check if a sheet with the respondent's name exists
  let individualSheet = worksheet.getSheetByName(respondentName);
  
  // If the sheet doesn't exist, create a new one
  if (!individualSheet) {
    individualSheet = worksheet.insertSheet(respondentName);
    let headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    individualSheet.appendRow(headers);
  }
  
  // Get the last row in the individual's sheet and determine the destination range
  let rowToPaste = individualSheet.getLastRow() + 1;
  let destinyRange = individualSheet.getRange(rowToPaste, 1, 1, columns);
  
  // Copy the response to the individual's sheet
  responseSheet.getRange(rowToCopy, 1, 1, columns).copyTo(destinyRange);
}
