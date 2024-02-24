/**
 * This function is triggered when the Google Apps Script is accessed as a web app.
 * It returns an HtmlOutput object, which is the HTML content of the web app.
 * @return {HtmlOutput} The HTML content of the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * This function processes the data from the Google Sheet and sends emails using the Google Word template.
 * @param {string} sheetUrl - The URL of the Google Sheet containing the data.
 * @param {string} templateUrl - The URL of the Google Word template.
 * @return {object} An object containing the status of the operation, any errors that occurred, and a message.
 */
function doPost(sheetUrl, templateUrl) {

  // Check if the sheetUrl and templateUrl are not empty
  if (!sheetUrl || !templateUrl) {
    return {
      status: false,
      message: "Sheet URL or Word URL is empty"
    };
  }

  // Open the Google Docs template and verify it's a Google Docs document
  var template = DocumentApp.openByUrl(templateUrl);
  if (!template.getBody) {
    return {
      status: false,
      message: "Invalid word template"
    };
  }

  // Check if the templateUrl is a valid Google Docs URL
  if (!templateUrl.startsWith('https://docs.google.com/document/d/')) {
    return {
      status: false,
      message: "Invalid Google Word URL found in excel sheet. Please use a valid Google Docs URL."
    };
  }
  

  // Access and process data from Google Sheet
  try {
    var sheet = SpreadsheetApp.openByUrl(sheetUrl).getSheetByName("Sheet1");
    var data = sheet.getDataRange().getValues();
  } catch (error) {
    return {
      status: false,
      message: "Unable to fetch the data from google sheet. Please allow the permission or make it public."
    };
  }

  // Process data row by row, sending emails with proper handling
  var emailsSent = 0;
  var errors = [];
  for (var i = 1; i < data.length; i++) {
    try {
      var receiverEmail = data[i][0];
      var receiverName = data[i][1];
      var emailSubject = data[i][2];

      let receiverEmailBody = template.getBody().getText().replace("{{name}}", receiverName);

      GmailApp.sendEmail(receiverEmail, emailSubject, receiverEmailBody, { htmlBody: receiverEmailBody });
      emailsSent++;

    } catch (error) {
      errors.push('Error sending email to ' + receiverEmail + ': ' + error.message);
    }
  }

  return {
    status: true,
    errors,
    message: `Total ${emailsSent} email sent.`
  };
}