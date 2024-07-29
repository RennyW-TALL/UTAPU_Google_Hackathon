function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('create New Docs','createNewGoogleDocs')
  menu.addToUi();
}

function sendEmails(){
    SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.
     .alert('You clicked the first menu item!');
}
function createNewGoogleDocs() {
  try {
    // Retrieve the spreadsheet and sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Sheet "Form Responses 1" not found!');
      Logger.log('Sheet "Form Responses 1" not found!');
      return;
    }
    Logger.log('Sheet "Form Responses 1" found.');

    // Get all the data from the sheet
    const rows = sheet.getDataRange().getValues();
    Logger.log('Sheet data retrieved.');

    // Loop through each row of data
    rows.forEach(function(row, index) {
      if (index === 0) return; // Skip the header row
      if (row[9]) return; // Skip rows that already have a URL

      // Check if required fields are empty
      if (!row[1] || !row[2] || !row[3] || !row[4] || !row[8]) {
        Logger.log(`Skipping row ${index + 1} due to missing data.`);
        return; // Continue to the next row
      }

      let folderId;

      // Determine the folder ID based on the department
      switch (row[4]) {
        case 'Finances':
          folderId = '1bqyjG1AIhyKUrIfdvsoJ39pjigRhdf1a'; // Folder ID for Finances
          break;
        case 'IT Services':
          folderId = '14v1QSM7hI1dAxYQHJ35OJawmQXXDYwEj'; // Folder ID for IT Services
          break;
        case 'HR':
          folderId = '1Me5ZQOMcq4AF-XCeFT7TKC9n82jRjLyO'; // Folder ID for HR
          break;
        default:
          Logger.log(`Skipping row ${index + 1} due to unrecognized department.`);
          return; // Continue to the next row
      }

      try {
        const googleDocTemplate = DriveApp.getFileById('1AJC3DHrrig0-kBRL8BRlB6dKn33y1Gr8XJdg3WvFs2I'); // Template ID
        Logger.log('Template ID retrieved.');

        const destinationFolder = DriveApp.getFolderById(folderId);
        Logger.log('Destination folder ID retrieved.');

        Logger.log(`Creating document for row ${index + 1}`);
        const copy = googleDocTemplate.makeCopy(`${row[1]}, ${row[0]} Employee Details`, destinationFolder);
        Logger.log(`Document copy created for row ${index + 1}`);

        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        // Replace placeholders in the document with actual data
        body.replaceText('{{Full Name in Identification Card}}', row[1]);
        body.replaceText('{{Date of Birth}}', row[2]);
        body.replaceText('{{Email Address}}', row[3]);
        body.replaceText('{{Assigned Department}}', row[4]);
        body.replaceText('{{Consolidated Job Roles}}', row[8]);

        // Save and close the document
        doc.saveAndClose();
        Logger.log(`Document created and saved: ${doc.getId()}`);

        // Get the URL of the new document
        const url = doc.getUrl();
        Logger.log(`Document URL: ${url}`);

        // Insert the URL into the sheet in the 10th column (index 9)
        sheet.getRange(index + 1, 10).setValue(url);
        Logger.log(`URL set in sheet for row ${index + 1}`);
      } catch (error) {
        Logger.log(`Error creating document for row ${index + 1}: ${error}`);
      }
    });

    Logger.log('Process completed.');
  } catch (error) {
    Logger.log(`Error in createNewGoogleDocs function: ${error}`);
    SpreadsheetApp.getUi().alert(`Process stopped: ${error.message}`);
  }
}
