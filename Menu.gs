function onOpen() {
  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Gmail to sheets')
    .addItem('Config','showSidebar')
    .addItem('Fetch Data', 'fetchData')
    .addToUi();
}

function showSidebar() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('menu')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Gmail to sheets')
    .setWidth(300)
    .setHeight(100);
  ui.showSidebar(htmlOutput);
}

function getGmailLabels() {
  var labels = GmailApp.getUserLabels();
  var labelNames = labels.map(function(label) {
    var originalName = label.getName()
    return originalName;
  });
  return labelNames;
  
}

function getSheetNames() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(function(sheet) {
    return sheet.getName();
  });
  return sheetNames;
}


function setGmailLabel(labelName) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(ConfigVars.GMAIL_LABEL, labelName);
  
}

function setSalesDataSheet(sheetName) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(ConfigVars.SHEET_NAME, sheetName);
}