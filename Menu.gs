function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 IO Automation')
    .addItem('1. Generate IO Sheets', 'generateIOs')
    .addItem('2. Export Sheets to PDF', 'printPDF')
    .addToUi();
}
