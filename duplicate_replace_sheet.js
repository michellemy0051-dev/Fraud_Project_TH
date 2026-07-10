function duplicate_sus_player_without_percentile_lt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Suspicious_Player");
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Suspicious_Player' sheet not found.");
    return;
  }

  // 1. Get the dates from the cells
  const fromDate = sourceSheet.getRange("B3").getValue();
  const toDate = sourceSheet.getRange("B4").getValue();
  
  if (!fromDate || !toDate) {
    SpreadsheetApp.getUi().alert("Error: Please make sure both B3 and B4 contain valid dates.");
    return;
  }

  // 2. Format the dates
  const fromDateFormatted = Utilities.formatDate(new Date(fromDate), ss.getSpreadsheetTimeZone(), "d MMMM").toLowerCase();
  const toDateFormatted = Utilities.formatDate(new Date(toDate), ss.getSpreadsheetTimeZone(), "d MMMM yyyy").toLowerCase();
  
  const newSheetName = `${fromDateFormatted} ~ ${toDateFormatted} sus user`;

  // 3. Check if the target sheet already exists
  let targetSheet = ss.getSheetByName(newSheetName);
  
  if (targetSheet) {
    targetSheet.clear();
    targetSheet.setFrozenRows(0); // Unfreeze before clearing to prevent layout locked errors
    Logger.log("Existing sheet found and cleared: " + newSheetName);
  } else {
    targetSheet = ss.insertSheet(newSheetName);
    Logger.log("Created new sheet: " + newSheetName);
  }
  
  const sourceLastRow = sourceSheet.getLastRow();
  const maxSourceColumns = sourceSheet.getLastColumn();
  
  // Fetch Column A values to search for our slice boundaries
  const colAValues = sourceSheet.getRange(1, 1, sourceLastRow, 1).getValues();
  
  let overallStatsRow = -1;
  let userInfoRow = -1;

  // Scan Column A for the vertical row markers
  for (let r = 0; r < colAValues.length; r++) {
    const cellValue = String(colAValues[r][0]).trim();
    if (cellValue === "Overall Stats") {
      overallStatsRow = r + 1; // 1-based index
    }
    if (cellValue === "User Info") {
      userInfoRow = r + 1; // 1-based index
      break; 
    }
  }

  // Fallback checks for row markers
  if (overallStatsRow === -1 || userInfoRow === -1) {
    SpreadsheetApp.getUi().alert("Error: Could not locate 'Overall Stats' or 'User Info' in Column A.");
    return;
  }

  // Look exactly 1 row AFTER "User Info" to find headers (Row 25)
  const headerRowIndex = userInfoRow + 1; 
  const headerValues = sourceSheet.getRange(headerRowIndex, 1, 1, maxSourceColumns).getValues()[0];
  let targetColumnsCount = -1;

  // Find the total boundary column (ltTurnover)
  for (let c = 0; c < headerValues.length; c++) {
    if (String(headerValues[c]).trim().toLowerCase() === "ltturnover") {
      targetColumnsCount = c + 1; 
      break;
    }
  }

  // Fallback: if "ltturnover" isn't found, copy all columns
  if (targetColumnsCount === -1) {
    targetColumnsCount = maxSourceColumns;
  }

  // Ensure target sheet has enough columns
  const targetMaxColumns = targetSheet.getMaxColumns();
  if (targetMaxColumns < targetColumnsCount) {
    targetSheet.insertColumnsAfter(targetMaxColumns, targetColumnsCount - targetMaxColumns);
  }

  // === EXECUTING THE TWO-PART COPY ===
  const topRowsCount = overallStatsRow - 1;
  if (topRowsCount > 0) {
    sourceSheet.getRange(1, 1, topRowsCount, targetColumnsCount).copyTo(targetSheet.getRange(1, 1));
  }

  const bottomRowsCount = sourceLastRow - userInfoRow + 1;
  const targetPasteRow = topRowsCount + 1;
  if (bottomRowsCount > 0) {
    sourceSheet.getRange(userInfoRow, 1, bottomRowsCount, targetColumnsCount).copyTo(targetSheet.getRange(targetPasteRow, 1));
  }
  
  // Copy column widths
  for (let i = 1; i <= targetColumnsCount; i++) {
    targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
  }

  // === NEW: DELETE THE UNWANTED COLUMNS FROM THE TARGET SHEET ===
  // We scan the newly pasted headers on the target sheet from right to left and drop unwanted ones
  const targetHeaders = targetSheet.getRange(targetPasteRow + 1, 1, 1, targetColumnsCount).getValues()[0];
  const columnsToHide = ["ltsuccessdpcnt", "ltfaildpcnt", "ltwdcnt", "ltclaimpromo"];

  for (let i = targetHeaders.length - 1; i >= 0; i--) {
    const headerName = String(targetHeaders[i]).trim().toLowerCase();
    if (columnsToHide.includes(headerName)) {
      targetSheet.deleteColumn(i + 1); // 1-based index
    }
  }

  // === FREEZE THE MAIN DATA HEADERS ROW ===
  const newHeaderRowIndex = targetPasteRow + 1; 
  targetSheet.setFrozenRows(newHeaderRowIndex);

  // --- Clean up empty rows at the bottom ---
  const lastRowWithData = targetSheet.getLastRow();
  const totalRowsNow = targetSheet.getMaxRows();
  const desiredTotalRows = lastRowWithData + 5;

  if (totalRowsNow > desiredTotalRows) {
    targetSheet.deleteRows(desiredTotalRows + 1, totalRowsNow - desiredTotalRows);
  } else if (totalRowsNow < desiredTotalRows) {
    targetSheet.insertRowsAfter(totalRowsNow, desiredTotalRows - totalRowsNow);
  }

  // Bring the user to the newly generated sheet
  ss.setActiveSheet(targetSheet);
}