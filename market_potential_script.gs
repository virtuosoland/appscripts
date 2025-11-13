/**
 * @OnlyCurrentDoc
 *
 * This script adds a custom menu to highlight and sort rows based on market potential.
 */

// Define the colors for highlighting
const EXPANSIVE_COLOR = '#a4c2f4'; // Light Blue
const MIDDLE_MARKET_COLOR = '#b6d7a8'; // Green
const LOWER_POTENTIAL_COLOR = '#f6b26b'; // Orange

// Define the category names
const EXPANSIVE_NAME = "1 - Expansive Market";
const MIDDLE_NAME = "2 - Medium Value Market";
const LOWER_NAME = "3 - Lower Volume Market";
const OTHER_NAME = ""; // Or "Other" if you prefer
const CATEGORY_HEADER = "Market Category";
const SORT_KEY_HEADER = "_SortKey"; // Helper column for sorting

/**
 * Creates a custom menu in the spreadsheet when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Market Analysis')
    .addItem('Process & Sort Sheet', 'sortPotentials')
    .addToUi();
}

/**
 * Ensures helper columns exist, adding them if not.
 * @param {Sheet} sheet The active Google Sheet.
 * @return {boolean} True if columns are ready, false if something failed.
 */
function ensureHelperColumns(sheet) {
  Logger.log("Running ensureHelperColumns...");
  const ui = SpreadsheetApp.getUi();
  const headerRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  let headers = headerRange.getValues()[0];
  let colAdded = false;

  const columnsToVerify = [CATEGORY_HEADER, SORT_KEY_HEADER];

  columnsToVerify.forEach(headerName => {
    if (headers.indexOf(headerName) === -1) {
      Logger.log(`Column "${headerName}" not found, attempting to add it.`);
      // Column not found, add it
      try {
        const newColIndex = sheet.getLastColumn() + 1;
        sheet.getRange(2, newColIndex).setValue(headerName);
        colAdded = true;
        Logger.log(`Column "${headerName}" added successfully.`);
      } catch (e) {
        Logger.log(`ERROR adding column "${headerName}": ${e.message}`);
        ui.alert(`Could not add "${headerName}" column. Error: ${e.message}. Please add it manually on row 2 and re-run.`);
        return false; // Stop this helper
      }
    }
  });

  if (colAdded) {
    Logger.log("Columns were added, flushing sheet.");
    SpreadsheetApp.flush(); // Ensure changes are committed
  }
  Logger.log("ensureHelperColumns finished.");
  return true; // All columns should be ready
}

/**
 * Re-fetches column indices after potential column addition.
 * @param {Sheet} sheet The active Google Sheet.
 * @return {object} An object mapping rule names to column indices, or null if headers aren't ready.
 */
function getColumnIndices(sheet) {
  Logger.log("Running getColumnIndices...");
  if (!ensureHelperColumns(sheet)) {
    Logger.log("ensureHelperColumns returned false.");
    return null; // Stop execution, error was shown
  }
  
  // Re-fetch headers *after* columns might have been added
  const headerRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  Logger.log("Headers found: " + headers.join(", "));
  let colIndices = findHeaderColumns(headers);
  Logger.log("Column indices found: " + JSON.stringify(colIndices));


  // Final check
  if (colIndices.marketCategory === -1 || colIndices.sortKey === -1) {
     Logger.log("ERROR: Helper columns not found after creation attempt.");
     SpreadsheetApp.getUi().alert(`Failed to find or create helper columns. Please manually add "${CATEGORY_HEADER}" and "${SORT_KEY_HEADER}" to row 2 and re-run.`);
     return null;
  }
  
  Logger.log("getColumnIndices finished successfully.");
  return colIndices;
}

/**
 * Sorts the sheet by potential and outputs to a NEW sheet.
 * THIS VERSION COPIES THE SHEET to preserve formatting, then sorts the copy.
 */
function sortPotentials() {
  Logger.log("--- Running sortPotentials ---");
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the whole workbook
  const activeSheet = spreadsheet.getActiveSheet(); // This is the SOURCE sheet
  
  if (!activeSheet) {
    Logger.log("Exiting: No active sheet found.");
    ui.alert("No active sheet found.");
    return;
  }
  
  const sourceSheetName = activeSheet.getName();
  const sourceSheetIndex = activeSheet.getIndex(); // Get the position of the original sheet
  Logger.log("Active sheet: " + sourceSheetName + " at index: " + sourceSheetIndex);

  // Check if it's already a processed sheet
  if (sourceSheetName.endsWith(" - processed")) {
    ui.alert("Please run this function on a raw data sheet, not on an already-processed sheet.");
    return;
  }

  const lastRow = activeSheet.getLastRow();
  const lastCol = activeSheet.getLastColumn();
  if (lastRow < 3) {
    Logger.log("Exiting: Not enough data in the sheet.");
    ui.alert("Not enough data in the sheet to sort. This script expects headers on row 2 and data starting on row 3.");
    return;
  }
  
  // Set the new sheet name
  const newSheetName = `${sourceSheetName} - processed`;

  ui.alert("Categorizing and sorting data... A new sheet named '" + newSheetName + "' will be created with the results. This will preserve all formatting.");

  // --- 1. Create the NEW Sheet by copying the active one ---
  Logger.log("Attempting to delete old processed sheet if it exists...");
  let sortedSheet = spreadsheet.getSheetByName(newSheetName);
  if (sortedSheet) {
    spreadsheet.deleteSheet(sortedSheet);
    SpreadsheetApp.flush(); // Make sure deletion is complete
    Logger.log("Old sheet deleted.");
  }
  
  Logger.log("Creating new sheet: " + newSheetName + " by copying.");
  sortedSheet = activeSheet.copyTo(spreadsheet).setName(newSheetName);
  
  // Move the new sheet to be right after the original one
  spreadsheet.setActiveSheet(sortedSheet);
  spreadsheet.moveActiveSheet(sourceSheetIndex + 1);
  
  SpreadsheetApp.flush();
  Logger.log("New sheet created, activated, and moved to index " + (sourceSheetIndex + 1));


  // --- 2. Get Column Indices from NEW sheet (and add helper columns) ---
  Logger.log("Getting column indices from new sheet...");
  // Note: getColumnIndices will now run on `sortedSheet`
  let colIndices = getColumnIndices(sortedSheet); 
  if (!colIndices) {
    Logger.log("Exiting: colIndices is null.");
    return; // Stop if ensureHelperColumns failed (it shows its own alert)
  }
  
  // We need to get the *new* last column in case helper columns were added
  const newLastCol = sortedSheet.getLastColumn();
  
  const missingCols = Object.keys(colIndices)
                           .filter(key => key !== 'marketCategory' && key !== 'sortKey' && colIndices[key] === -1);
                           
  if (missingCols.length > 0) {
    Logger.log("Exiting: Missing required columns: " + missingCols.join(", "));
    ui.alert(`Error: Could not find the following required columns: ${missingCols.join(", ")}. Please check the header names on Row 2.`);
    return;
  }
  
  // --- 3. Read data, calculate categories, SORT IN MEMORY, and WRITE back ---
  Logger.log("Reading data to calculate and sort in memory...");
  // Get the range from row 3 to the last data row
  const dataRegion = sortedSheet.getRange(3, 1, lastRow - 2, newLastCol);
  const dataValues = dataRegion.getValues();
  Logger.log(`Read ${dataValues.length} data rows.`);
  
  // Loop 1: Calculate categories and keys
  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    
    const sellThroughRate = Number(row[colIndices.sellThrough]);
    const totalSold = Number(row[colIndices.totalSold]);
    const percentOver200k = Number(row[colIndices.percentOver200k]);
    const percentUnder10k = Number(row[colIndices.percentUnder10k]);

    let categoryName = OTHER_NAME;
    let categoryKey = 4; // 4 = Other (default)

    if (!isNaN(sellThroughRate) && !isNaN(totalSold) && !isNaN(percentOver200k) && !isNaN(percentUnder10k)) {
      // Rule 1: Expansive Market
      if (sellThroughRate > 0.9 && totalSold > 100 && percentOver200k >= 0.4) {
        categoryName = EXPANSIVE_NAME;
        categoryKey = 1;
      } 
      // Rule 2: Middle Market
      else if (sellThroughRate > 0.9 && totalSold > 150 && percentUnder10k <= 0.1) {
        categoryName = MIDDLE_NAME;
        categoryKey = 2;
      }
      // Rule 3: Lower Potential
      else if (sellThroughRate > 0.8 && totalSold > 60 && percentUnder10k <= 0.1) {
        categoryName = LOWER_NAME;
        categoryKey = 3;
      }
    }
    
    // Write the category and key *into the data array*
    row[colIndices.marketCategory] = categoryName;
    row[colIndices.sortKey] = categoryKey;
  }
  
  // Loop 2: Sort the data array in memory
  Logger.log("Sorting data in memory...");
  dataValues.sort((a, b) => {
    const keyA = a[colIndices.sortKey];
    const keyB = b[colIndices.sortKey];
    
    // Primary sort: Category Key (1, 2, 3...)
    if (keyA !== keyB) {
      return keyA - keyB;
    }
    
    // Secondary sort: Sell Through Rate (High to Low)
    const sellThroughA = a[colIndices.sellThrough];
    const sellThroughB = b[colIndices.sellThrough];
    return sellThroughB - sellThroughA;
  });
  Logger.log("In-memory sort complete.");

  // Loop 3: Write the sorted data back to the sheet
  Logger.log("Writing sorted data back to sheet...");
  dataRegion.setValues(dataValues);
  SpreadsheetApp.flush();
  Logger.log("Sorted data written.");


  // --- 4. Apply Highlights on NEW Sheet ---
  Logger.log("Applying highlights to new sheet...");
  
  // We already have the sorted data, just need to build range lists
  let expansiveMarketRanges = [];
  let middleMarketRanges = [];
  let lowerPotentialRanges = [];

  for (let i = 0; i < dataValues.length; i++) {
    const rowNum = i + 3; // i=0 is Row 3
    const category = dataValues[i][colIndices.sortKey];
    const rowRangeA1 = sortedSheet.getRange(rowNum, 1, 1, newLastCol).getA1Notation(); 

    if (category === 1) expansiveMarketRanges.push(rowRangeA1);
    else if (category === 2) middleMarketRanges.push(rowRangeA1);
    else if (category === 3) lowerPotentialRanges.push(rowRangeA1);
  }

  // Clear all old backgrounds (on new sheet) - this is safe, formats are preserved
  dataRegion.setBackground(null);
  
  // Apply formatting in batches (on new sheet)
  if (expansiveMarketRanges.length > 0) sortedSheet.getRangeList(expansiveMarketRanges).setBackground(EXPANSIVE_COLOR);
  if (middleMarketRanges.length > 0) sortedSheet.getRangeList(middleMarketRanges).setBackground(MIDDLE_MARKET_COLOR);
  if (lowerPotentialRanges.length > 0) sortedSheet.getRangeList(lowerPotentialRanges).setBackground(LOWER_POTENTIAL_COLOR);
  Logger.log("Highlighting complete.");


  // --- 5. Format the NEW Sheet ---
  Logger.log("Formatting new sheet (freezing panes, hiding column)...");
  try {
    // Remove any filters that might have been copied
    const filter = sortedSheet.getFilter();
    if (filter) {
      filter.remove();
    }
    // Unfreeze all panes (in case of odd copy behavior)
    sortedSheet.unfreeze();
    SpreadsheetApp.flush();

    // Set the frozen panes as requested
    sortedSheet.setFrozenRows(2);
    sortedSheet.setFrozenColumns(1);

    // Hide the helper sort key column
    sortedSheet.hideColumns(colIndices.sortKey + 1);
  } catch (e) {
     Logger.log("Error formatting new sheet: " + e.message);
  }
  
  // --- 6. Show final alert ---
  Logger.log("--- sortPotentials finished ---");
  ui.alert(`Processing complete! \n\nA new sheet named '${newSheetName}' has been created with all original formatting preserved.`);
}


/**
 * Helper function to find the column index for each required header.
 * @param {Array<string>} headers An array of header names (from row 2).
 * @return {object} An object mapping rule names to column indices.
 */
function findHeaderColumns(headers) {
  return {
    sellThrough: headers.indexOf("Sold/OnMarket (Sell Through Rate) - Bigger the better"),
    totalSold: headers.indexOf("Total # of Sold Properties in 6 months"),
    percentOver200k: headers.indexOf("% (>200k)"),
    percentUnder10k: headers.indexOf("%<10k"),
    marketCategory: headers.indexOf(CATEGORY_HEADER),
    sortKey: headers.indexOf(SORT_KEY_HEADER)
  };
}