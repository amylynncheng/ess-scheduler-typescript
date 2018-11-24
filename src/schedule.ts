/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Helper')
    .addItem('Generate schedule', 'generateSchedule')
    .addToUi();
}

function generateSchedule() {
}