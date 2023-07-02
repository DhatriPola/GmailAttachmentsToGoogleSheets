function saveXlsxDataToSheets() {
  var documentProperties = PropertiesService.getDocumentProperties();
  const label = documentProperties.getProperty(ConfigVars.GMAIL_LABEL);
  var threads = GmailApp.getUserLabelByName(label).getThreads();
  var salesDataSheetName = documentProperties.getProperty(ConfigVars.SHEET_NAME);

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();

      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        
        // Check if the attachment is an XLSX file
        if (attachment.getContentType() === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
          var file = Drive.Files.insert({ title : 'temp_converted'}, attachment, {
        convert: true
      });
          var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
          var sheet = spreadsheet.getSheetByName(salesDataSheetName);
          
          const sourceValues = getExcelData(file.id)
          Logger.log(sourceValues)
          // Get the XLSX file as an Excel blob
          
          sheet.getRange(sheet.getLastRow()+1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
          
          // Delete the temporary XLSX file from Google Drive
          DriveApp.getFileById(file.getId()).setTrashed(true);
          
          // Mark the email as read or perform any other desired actions
          message.markRead();
          
          // Break the loop to process one XLSX file per email
          break;
        }
      }
    }
  }
}


function getExcelData(tempID) { 
  const source = SpreadsheetApp.openById(tempID);
  //The sheetname of the excel where you want the data from
  const sourceSheet = source.getSheets()[0];
  //The range you want the data from

  var lastRow = sourceSheet.getLastRow();
  var lastColumn = sourceSheet.getLastColumn();
  return sourceSheet.getRange(1,1,lastRow, lastColumn).getValues();
}

function fetchData() {
  var documentProperties = PropertiesService.getDocumentProperties();
  const labelName = documentProperties.getProperty(ConfigVars.GMAIL_LABEL);
  const sheetName = documentProperties.getProperty(ConfigVars.SHEET_NAME);

  var htmlOutput = HtmlService.createHtmlOutput(labelName + " - " + sheetName)
    .setWidth(300)
    .setHeight(100);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  saveXlsxDataToSheets()
  }
