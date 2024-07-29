function onFormSubmit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var email = range.getCell(1, 4).getValue(); // Replace <COLUMN_INDEX> with the column index of the email address
  
  var subject = "Thank you for your submission";
  var message = "Dear new hire,\n\nThank you for submitting your form. We have received your details and will process them shortly.\n\nBest regards,\nCompany GWH";
  
  MailApp.sendEmail(email, subject, message);
}

function setupTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmit')
           .forSpreadsheet(sheet)
           .onFormSubmit()
           .create();
}
