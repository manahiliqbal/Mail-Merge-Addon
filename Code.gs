function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Open Mail Merge', 'openMailMergeDialog')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function openMailMergeDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(600)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mail Merge');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getHtmlContent(filename) {
  return include(filename);
}

function getTemplateList() {
  return [
    { id: 'template1', name: 'Template 1' },
    { id: 'template2', name: 'Template 2' }
  ];
}

function getSheetList() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => ({ name: sheet.getName() }));
}

function scheduleEmails(scheduledDate, scheduledTime, timezone, emailInterval) {
  Logger.log('Scheduled Date: ' + scheduledDate);
  Logger.log('Scheduled Time: ' + scheduledTime);
  Logger.log('Timezone: ' + timezone);
  Logger.log('Email Interval: ' + emailInterval);
}
