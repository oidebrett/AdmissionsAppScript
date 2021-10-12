function onChange(e){    
  //Set up the dynamic cell and col references
  var configRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("A:G");
  var SendEmailSheet = configRange.getCell(37,3).getValue();
  var SendEmailCell = configRange.getCell(37,7).getValue();
  var thisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SendEmailSheet);
  var sendEmail = thisSheet.getRange(SendEmailCell).getValue();
  
  myObject = admissionsObject({
    sendEmails: sendEmail,
  })  
  myObject.run(e);
}

function sendEmailsPeriodically(){
  //Set up the dynamic cell and col references
  var configRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("A:G");
  var SendEmailSheet = configRange.getCell(37,3).getValue();
  var SendEmailCell = configRange.getCell(37,7).getValue();
  var thisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SendEmailSheet);
  var sendEmail = thisSheet.getRange(SendEmailCell).getValue();
  
  myObject = admissionsObject({
    sendEmails: sendEmail,
  })  
  myObject.sendEmailsPeriodically();
}
