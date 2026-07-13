/**
 * 1. GENERATE IOS
 * Loops through "Manual Top Up Request" sheet, checking Column U for "For IO Creation"
 * and populates the entire template.
 * Updated to include protective string formatting and skip logging.
 */
function generateIOs() {
  if (typeof checkPermission === 'function' && !checkPermission()) return;
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Manual Top Up Request");
  const templateSheet = ss.getSheetByName("_template_");
 
  if (!sourceSheet || !templateSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Manual Top Up Request' or '_template_' sheet not found.");
    return;
  }
 
  const headerRowIndex = 2;
  const startRowIndex = 3;
  const lastRow = sourceSheet.getLastRow();
 
  if (lastRow < startRowIndex) {
    SpreadsheetApp.getUi().alert("No data found in 'Manual Top Up Request'.");
    return;
  }
 
  const totalColumns = sourceSheet.getLastColumn();
 
  // Find the IO Number Column dynamically by reading Row 2 headers
  const headers = sourceSheet.getRange(headerRowIndex, 1, 1, totalColumns).getValues()[0];
  let ioColumnIndex = -1;
 
  for (let i = 0; i < headers.length; i++) {
    const headerText = headers[i].toString().toLowerCase();
    if (headerText.indexOf("insertion order") !== -1 || headerText.indexOf("io number") !== -1 || headerText.indexOf("io no") !== -1) {
      ioColumnIndex = i + 1;
      break;
    }
  }
 
  if (ioColumnIndex === -1) {
    ioColumnIndex = 19; // Column S default
  }
 
  const statusColumnIndex = 21; // Column U
 
  const numRowsToFetch = lastRow - startRowIndex + 1;
  
  // Fetch all row data at once to optimize performance and prevent API timeout
  const allRowData = sourceSheet.getRange(startRowIndex, 1, numRowsToFetch, totalColumns).getValues();
  let generatedCount = 0;
  let skippedRows = []; // Tracks rows that matched "For IO Creation" but missed an IO Number
 
  for (let i = 0; i < allRowData.length; i++) {
    const row = allRowData[i];
    const currentRowIndex = startRowIndex + i;
    
    const ioValue = row[ioColumnIndex - 1] ? row[ioColumnIndex - 1].toString().trim() : "";
    const statusValue = row[statusColumnIndex - 1] ? row[statusColumnIndex - 1].toString().trim() : "";
   
    // Check for status match (Case-Insensitive)
    if (statusValue.toLowerCase() === "for io creation") {
      
      // If status is correct but IO value is blank, track it and skip
      if (ioValue === "") {
        skippedRows.push(currentRowIndex);
        continue;
      }
      
      // Extract details from the current row mapping
      const shopName = row[5];          // Column F (Index 5)
      const effectiveDate = row[7];     // Column H (Index 7)
      const kamName = row[8];           // Column I (Index 8)
      const customerAddress = row[10];  // Column K (Index 10)
      const sbcCustomerName = row[11];  // Column L (Index 11)
      const customerEmail = row[12];    // Column M (Index 12)
      
      // Create a fresh copy of the template
      const newSheet = templateSheet.copyTo(ss);
      
      // --- POPULATE THE ENTIRE TEMPLATE ---
      newSheet.getRange("O11").setValue(ioValue);          // IO Number
      newSheet.getRange("O12").setValue(shopName);         // Shop Name
      newSheet.getRange("I19").setValue(sbcCustomerName);  // SBC Customer Name / Attn
      newSheet.getRange("I20").setValue(customerEmail);    // Customer Email
      newSheet.getRange("D22").setValue(kamName);          // Account Manager / KAM Name
      newSheet.getRange("D23").setValue(effectiveDate);    // Effective Date
      
      // Handle naming and sanitization
      let sheetName = shopName ? shopName.toString() : "IO Sheet";
      sheetName = cleanFilename(sheetName, "-");
      
      if (sheetName.length > 100) sheetName = sheetName.substring(0, 100);
      
      let finalName = sheetName;
      let count = 1;
      while (ss.getSheetByName(finalName)) {
        finalName = sheetName + " (" + count + ")";
        count++;
      }
      
      newSheet.setName(finalName);
      
      // Update status back on the main tracking sheet
      sourceSheet.getRange(currentRowIndex, statusColumnIndex).setValue("Done IO Creation");
      
      generatedCount++;
    }
  }
 
  // UI Alerts Handling
  if (generatedCount === 0) {
    let msg = "No rows found matching 'For IO Creation' in column U.";
    if (skippedRows.length > 0) {
      msg += `\n\nNote: Rows [${skippedRows.join(", ")}] had the right status but were skipped because their IO Number column was completely empty 🦆.`;
    }
    SpreadsheetApp.getUi().alert(msg);
  } else {
    let msg = `Successfully generated ${generatedCount} IO Sheets with complete information and updated their status to 'Done IO Creation'!`;
    if (skippedRows.length > 0) {
      msg += `\n\n⚠️ Warning: Rows [${skippedRows.join(", ")}] were skipped because they were missing an IO Number.`;
    }
    msg += "\n\nProceed to Step 2 to export 🦆.";
    SpreadsheetApp.getUi().alert(msg);
  }
}

/**
 * Helper function to sanitize sheet names if missing from the core script
 */
function cleanFilename(name, replaceChar) {
  if (!name) return "Untitled";
  return name.toString().replace(/[\\\/\?\*\:\[\]]/g, replaceChar || "-").trim();
}
