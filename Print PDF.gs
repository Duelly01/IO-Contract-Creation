/**
 * 2. PRINT PDF (Optimized for My Ads Credit Purchase Agreement Layout)
 * Dynamically routes to folder paths, matches temporary tabs back to source data,
 * forces the export to skip the top rows (master data dumps) so that the contract
 * begins immediately on Page 1, and deletes the temporary sheet tab upon completion.
 * Updated: Extended row boundaries from 125 to 245 to include all 3 pages.
 */
function printPDF() {
  if (!checkPermission()) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Manual Top Up Request");
  const sheets = ss.getSheets();
  const excludedSheets = ["Manual Top Up Request", "_template_", "INFO", "Seller's Details", "IO Data", "Form Responses 1", "Template Manual Top Ups", "Free Ads Credit Request", "Special Payment Term Brands", "Adhoc Template Manual Top Ups"];
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Manual Top Up Request' sheet not found.");
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  const startRowIndex = 3;
  let ioNoMapping = {};
  if (lastRow >= startRowIndex) {
    const totalColumns = sourceSheet.getLastColumn();
    const headers = sourceSheet.getRange(2, 1, 1, totalColumns).getValues()[0];
    let ioColumnIndex = -1;
    let folderColumnIndex = -1;
    
    // Find both column indices dynamically to prevent layout tracking issues
    for (let i = 0; i < headers.length; i++) {
      const headerText = headers[i].toString().toLowerCase();
      if (ioColumnIndex === -1 && (headerText.indexOf("insertion order") !== -1 || headerText.indexOf("io number") !== -1 || headerText.indexOf("io no") !== -1)) {
        ioColumnIndex = i + 1;
      }
      if (folderColumnIndex === -1 && (headerText.indexOf("folder") !== -1 || headerText.indexOf("IO Link") !== -1)) {
        folderColumnIndex = i + 1;
      }
    }
    
    // Fallbacks if headers can't be matched by keywords
    if (ioColumnIndex === -1) ioColumnIndex = 19; // Column S
    if (folderColumnIndex === -1) folderColumnIndex = 20; // Column T

    const maxColumnsToFetch = Math.max(ioColumnIndex, folderColumnIndex);
    const sourceData = sourceSheet.getRange(startRowIndex, 1, lastRow - startRowIndex + 1, maxColumnsToFetch).getValues();
    
    let chipUrls = [];
    try {
      const sheetName = sourceSheet.getName();
      const columnLetter = sourceSheet.getRange(1, folderColumnIndex).getA1Notation().replace(/\d+/g, '');
      const rangeA1 = `'${sheetName}'!${columnLetter}${startRowIndex}:${columnLetter}${lastRow}`;
      
      const apiResponse = Sheets.Spreadsheets.get(ss.getId(), {
        ranges: [rangeA1],
        fields: "sheets(data(rowData(values(userEnteredValue,chipRuns))))"
      });
      const rowData = apiResponse.sheets[0].data[0].rowData;
      if (rowData) {
        chipUrls = rowData.map(row => {
          if (row.values && row.values[0] && row.values[0].chipRuns && row.values[0].chipRuns[0] && row.values[0].chipRuns[0].chip) {
            return row.values[0].chipRuns[0].chip.richLinkProperties.uri || "";
          }
          return "";
        });
      }
    } catch (apiError) {
      Logger.log(`[CHIP EXTRACTION ERROR] Failed to fetch smart chips via Advanced API: ${apiError.message}. Falling back to plain text extraction.`);
    }

    for (let r = 0; r < sourceData.length; r++) {
      const ioKey = sourceData[r][ioColumnIndex - 1].toString().trim();
      let folderVal = (chipUrls[r] && chipUrls[r] !== "") ? chipUrls[r] : "";
      if (!folderVal && sourceData[r][folderColumnIndex - 1]) {
        folderVal = sourceData[r][folderColumnIndex - 1].toString().trim();
      }
      
      if (ioKey !== "") {
        ioNoMapping[ioKey] = folderVal;
      }
    }
  }

  let exportCount = 0;
  
  // Loop backwards through sheets to safely process and delete them
  for (let i = sheets.length - 1; i >= 0; i--) {
    const ws = sheets[i];
    const sheetName = ws.getName();
    
    if (excludedSheets.indexOf(sheetName) === -1) {
      const ioNumberRaw = ws.getRange("O11").getValue().toString();
      const ioNumber = ioNumberRaw.trim();
      
      if (!ioNumber) {
        Logger.log(`Skipping tab ${sheetName} because no Insertion Order number was found in cell O11.`);
        continue;
      }

      const cleanIoNumber = cleanFilename(ioNumber, "-");
      const fileName = cleanIoNumber + ".pdf";
      let destinationFolder = DriveApp.getRootFolder(); 
      const customFolderValue = ioNoMapping[ioNumber];
      
      if (customFolderValue) {
        try {
          let folderId = customFolderValue;
          if (customFolderValue.indexOf("folders/") !== -1) {
            const match = customFolderValue.match(/\/folders\/([a-zA-Z0-9-_]+)/);
            if (match && match[1]) folderId = match[1];
          } else if (customFolderValue.indexOf("id=") !== -1) {
            const match = customFolderValue.match(/[?&]id=([a-zA-Z0-9-_]+)/);
            if (match && match[1]) folderId = match[1];
          }
          destinationFolder = DriveApp.getFolderById(folderId);
        } catch(e) {
          Logger.log(`[ROUTING NOTICE] Failed to open target folder for IO ${ioNumber}. Saving to My Drive (Root Folder) instead. Reason: ${e.message}`);
        }
      } else {
        Logger.log(`[ROUTING NOTICE] No folder value mapped for IO ${ioNumber}. Defaulting file destination to My Drive (Root Folder).`);
      }
      
      // ✂️ Clean up trailing formatting blocks beyond the contract grid bounds
      if (ws.getLastColumn() > 16) {
        ws.deleteColumns(17, ws.getLastColumn() - 16);
      }
      
      // 🛠️ FIX: Extended check from 125 to 245 to prevent truncating the 3rd page content
      if (ws.getLastRow() > 245) {
        ws.deleteRows(246, ws.getLastRow() - 245);
      }

      SpreadsheetApp.flush(); 
      
      const url = ss.getUrl().replace(/edit$/, '') + 'export?';

      // 🎯 PARAMETERS CONFIGURED TO CROP AWAY TOP ROWS 1-3
      const exportOptions = {
        exportFormat: 'pdf',
        format: 'pdf',
        size: 'letter',               // Letter Size
        portrait: 'true',             // Portrait Orientation
        fitw: 'true',                 // Fit Width to Page
        gridlines: 'false',           // Hide spreadsheet gridlines
        printtitle: 'false',          // Hide Sheet Title
        sheetnames: 'false',          // Hide Sheet Tab Name
        fzr: 'false',                 // Do not repeat frozen rows
        fzc: 'false',                 // Do not repeat frozen columns
        top_margin: '0.40',           // Tight top margin
        bottom_margin: '0.40',        // Tight bottom margin
        left_margin: '0.45',          // Balanced side margins
        right_margin: '0.45',
        gid: ws.getSheetId(),         
        
        // 🔒 CROPPING TARGET RANGE MATRIX
        r1: '3',                      // ⭐ SKIP TOP ROWS: Start at Row 4 (0-Indexed = 3) to drop master tables
        c1: '0',                      // Start at Column A (Index 0)
        r2: '245',                    // 🛠️ FIX: Extended End row boundary to include up to Row 245
        c2: '16'                      // Strict End column boundary (Column P)
      };
      
      let urlParts = [];
      for (let key in exportOptions) {
        urlParts.push(key + '=' + exportOptions[key]);
      }
      const exportUrl = url + urlParts.join('&');
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
          'MuteHttpExceptions': true
        }
      });
      const blob = response.getBlob().setName(fileName);
      
      // 📝 LOGGING
      try {
        Logger.log(`[EXPORTING] File: ${fileName} | Skipped rows 1-3 | Viewport set to Row 4 down to Row 245 | Target Folder: "${destinationFolder.getName()}"`);
      } catch (err) {
        Logger.log(`[EXPORTING] File: ${fileName} | Could not resolve folder metadata: ${err.message}`);
      }

      destinationFolder.createFile(blob);
      
      // 🗑️ DELETE THE TAB AFTER GENERATING THE PDF SUCCESSFULLY
      try {
        ss.deleteSheet(ws);
        Logger.log(`[Sweep Sweep Sweep] Deleted temporary tab: ${sheetName}`);
      } catch (deleteError) {
        Logger.log(`[SWEEP SWEEP SWEEP ERROR] Failed to delete sheet ${sheetName}: ${deleteError.message}`);
      }
      
      exportCount++;
    }
  }
  
  if (exportCount === 0) {
    SpreadsheetApp.getUi().alert("No temporary IO sheets found to export.");
  } else {
    SpreadsheetApp.getUi().alert(`Successfully exported ${exportCount} sheet(s) to their designated Google Drive folders and cleaned up the active tabs! 🦆`);
  }
}
