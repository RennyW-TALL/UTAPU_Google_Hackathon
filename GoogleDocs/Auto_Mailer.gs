function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('Auto Mailer');
  menu.addItem('Send Email', 'createNewGoogleDocs');
  menu.addToUi();
}

function createNewGoogleDocs() {
  var doc = DocumentApp.getActiveDocument().getBody().getTables();
  Logger.log('Number of tables: ' + doc.length);

  var arr = [];
  doc.forEach(table => {
    Logger.log('Number of rows in table: ' + table.getNumRows());

    // Ensure there are at least two rows in the table
    if (table.getNumRows() < 2) {
      Logger.log('Table does not have enough rows.');
      return;
    }

    var data = [];
    for (var i = 0; i < table.getNumRows(); i += 2) {
      var key = table.getRow(i).getCell(0).getText();
      var value = (i + 1 < table.getNumRows()) ? table.getRow(i + 1).getCell(0).getText() : 'No value';
      data.push({ key: key, value: value });
    }

    arr.push(data);
  });

  if (arr.length > 1 && arr[0].length > 2) {
    Logger.log(arr[0][2].value.toString());

    var subject = "Offer Letter";
    var message = "Dear " + arr[0][0].value + ",\n\nWe are pleased to inform you that your role as " + arr[1][1].value + " at our " + arr[1][0].value + " department has been approved.\n\nYou will be liaising with " + arr[1][3].value + " (your supervisor/manager).\n\nPlease feel free to ask any questions that you may have!\n\nBest regards,\nHR Manager\n123, Jalan Indah, 45012 Selangor";
    var htmlMessage = '<div style="font-family: Arial, sans-serif; background-color: #f4f4f4; color: #333; padding: 20px;">' +
                      '<div style="background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); max-width: 600px; margin: 0 auto;">' +
                      '<h1 style="color: #4CAF50;">Offer Letter</h1>' +
                      '<p>Dear ' + arr[0][0].value + ',</p>' +
                      '<p>We are pleased to inform you that your role as <strong>' + arr[1][1].value + '</strong> at our <strong>' + arr[1][0].value + '</strong> department has been approved.</p>' +
                      '<p>You will be liaising with <strong>' + arr[1][3].value + '</strong> (your supervisor/manager).</p>' +
                      '<p>Please feel free to ask any questions that you may have!</p>' +
                      '<p>Best regards,</p>' +
                      '<p>HR Manager<br>' +
                      '<span style="color: #777; font-size: 0.9em; margin-top: 10px;">123, Jalan Indah, 45012 Selangor</span></p>' +
                      '</div>' +
                      '</div>';

    MailApp.sendEmail(arr[0][2].value, subject, message, {
      htmlBody: htmlMessage
    });
  } else {
    Logger.log('Not enough data to send email.');
  }
}


