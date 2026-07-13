function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1"); // change if needed
  const lastRow = sheet.getLastRow();
  const amountCell = sheet.getRange(lastRow, 7); // Column G is 7
  amountCell.setNumberFormat("#,##0.00");
}
