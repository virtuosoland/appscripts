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
const SORTED_SHEET_NAME = "Sorted Market Analysis"; // Define the output sheet name

/**
 * Creates a custom menu in the spreadsheet when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Market Analysis')
    .addItem('Highlight Potentials (On This Sheet)', 'highlightPotentials')
    .addItem('Sort & Highlight (To New Sheet)', 'sortPotentials')
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
 * Main function to find and highlight high and medium potential rows.
 * This function just highlights; it does not sort.
 */
function highlightPotentials() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (!sheet) {
    ui.alert("No active sheet found.");
    return;
  }
  
  if (sheet.getName() === SORTED_SHEET_NAME) {
    ui.alert("This function highlights the raw data sheet. The sorted sheet is already highlighted.");
    return;
  }

  // Get data starting from Row 2 (headers)
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert("Not enough data in the sheet to process. This script expects headers on row 2 and data starting on row 3.");
    return;
  }

  let colIndices = getColumnIndices(sheet);
  if (!colIndices) {
    return; // Stop if ensureHelperColumns failed
  }
  
  // Check if all *other* required columns were found
  const missingCols = Object.keys(colIndices)
                           .filter(key => key !== 'marketCategory' && key !== 'sortKey' && colIndices[key] === -1);
                           
  if (missingCols.length > 0) {
    ui.alert(`Error: Could not find the following required columns: ${missingCols.join(", ")}. Please check the header names on Row 2.`);
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues(); // values[0] is headers, values[1] is row 3

  let expansiveMarketRanges = [];
  let middleMarketRanges = [];
  let lowerPotentialRanges = [];
  const categoryValues = []; // Array to hold the category names for batch writing
  const sortKeyValues = []; // Array to hold the sort keys
  const numCols = sheet.getLastColumn();

  // Define the range for data rows (excluding headers)
  const dataRowsRange = sheet.getRange(3, 1, values.length - 1, numCols);
  // Clear existing background colors first
  dataRowsRange.setBackground(null);

  // Start loop from row 3 (data starts at index 1 of 'values' array)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Parse values
    const sellThroughRate = Number(row[colIndices.sellThrough]);
    const totalSold = Number(row[colIndices.totalSold]);
    const percentOver200k = Number(row[colIndices.percentOver200k]);
    const percentUnder10k = Number(row[colIndices.percentUnder10k]);
    
    let categoryName = OTHER_NAME;
    let categoryKey = 4; // 4 = Other (default)

    // Skip rows with invalid data
    if (!isNaN(sellThroughRate) && !isNaN(totalSold) && !isNaN(percentOver200k) && !isNaN(percentUnder10k)) {
      const rowNum = i + 2; // +2 because values index 1 is Row 3
      const rowRangeA1 = sheet.getRange(rowNum, 1, 1, numCols).getA1Notation();

      // Rule 1: Expansive Market
      if (sellThroughRate > 0.9 && totalSold > 100 && percentOver200k >= 0.4) {
        expansiveMarketRanges.push(rowRangeA1);
        categoryName = EXPANSIVE_NAME;
        categoryKey = 1;
      } 
      // Rule 2: Middle Market
      else if (sellThroughRate > 0.9 && totalSold > 150 && percentUnder10k <= 0.1) {
        middleMarketRanges.push(rowRangeA1);
        categoryName = MIDDLE_NAME;
        categoryKey = 2;
      }
      // Rule 3: Lower Potential
      else if (sellThroughRate > 0.8 && totalSold > 60 && percentUnder10k <= 0.1) {
        lowerPotentialRanges.push(rowRangeA1);
        categoryName = LOWER_NAME;
        categoryKey = 3;
      }
    }
    
    categoryValues.push([categoryName]); // Add the category name for this row
    sortKeyValues.push([categoryKey]); // Add the sort key for this row
  }

  // Apply formatting in batches
  if (expansiveMarketRanges.length > 0) {
    sheet.getRangeList(expansiveMarketRanges).setBackground(EXPANSIVE_COLOR);
  }
  if (middleMarketRanges.length > 0) {
    sheet.getRangeList(middleMarketRanges).setBackground(MIDDLE_MARKET_COLOR);
  }
  if (lowerPotentialRanges.length > 0) {
    sheet.getRangeList(lowerPotentialRanges).setBackground(LOWER_POTENTIAL_COLOR);
  }
  
  // Write all category names and sort keys to their columns
  if (categoryValues.length > 0) {
    sheet.getRange(3, colIndices.marketCategory + 1, categoryValues.length, 1).setValues(categoryValues);
    sheet.getRange(3, colIndices.sortKey + 1, sortKeyValues.length, 1).setValues(sortKeyValues);
  }

  ui.alert(`Highlighting complete! Found ${expansiveMarketRanges.length} Expansive, ${middleMarketRanges.length} Medium Value, and ${lowerPotentialRanges.length} Lower Volume rows.`);
}

/**
 * Sorts the sheet by potential and outputs to a NEW sheet.
 * THIS VERSION SORTS IN JAVASCRIPT AND WRITES TO A NEW SHEET.
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
  Logger.log("Active sheet: " + activeSheet.getName());

  const lastRow = activeSheet.getLastRow();
  const lastCol = activeSheet.getLastColumn();
  if (lastRow < 3) {
    Logger.log("Exiting: Not enough data in the sheet.");
    ui.alert("Not enough data in the sheet to sort. This script expects headers on row 2 and data starting on row 3.");
    return;
  }

  ui.alert("Categorizing and sorting data... A new sheet named '" + SORTED_SHEET_NAME + "' will be created with the results.");

  // --- 1. Get Column Indices from SOURCE sheet ---
  Logger.log("Getting column indices...");
  let colIndices = getColumnIndices(activeSheet); 
  if (!colIndices) {
    Logger.log("Exiting: colIndices is null.");
    return; // Stop if ensureHelperColumns failed (it shows its own alert)
  }
  Logger.log("Column indices: " + JSON.stringify(colIndices));

  
  const missingCols = Object.keys(colIndices)
                           .filter(key => key !== 'marketCategory' && key !== 'sortKey' && colIndices[key] === -1);
                           
  if (missingCols.length > 0) {
    Logger.log("Exiting: Missing required columns: " + missingCols.join(", "));
    ui.alert(`Error: Could not find the following required columns: ${missingCols.join(", ")}. Please check the header names on Row 2.`);
    return;
  }
  
  // --- 2. Read Headers and Data from SOURCE sheet ---
  Logger.log("Reading header and data values...");
  const headerValues = activeSheet.getRange(1, 1, 2, lastCol).getValues();
  const dataValues = activeSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  Logger.log(`Read ${headerValues.length} header rows and ${dataValues.length} data rows.`);
  
  let sortableData = []; // Array to hold objects for sorting

  // --- 3. Calculate and Sort *in memory* (same logic as before) ---
  Logger.log("Categorizing and sorting in memory...");
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
    
    // Update the row *in memory* with the new category data
    row[colIndices.marketCategory] = categoryName;
    row[colIndices.sortKey] = categoryKey;
    
    // Add to our sortable array
    sortableData.push({
      row: row,
      sortKey: categoryKey,
      sellThrough: sellThroughRate
    });
  }

  // Sort the data *in memory*
  sortableData.sort((a, b) => {
    if (a.sortKey !== b.sortKey) {
      return a.sortKey - b.sortKey; // Ascending by category (1, 2, 3)
    }
    // If keys are the same, sort by sell-through descending
    return b.sellThrough - a.sellThrough;
  });

  // Extract the sorted row data
  const sortedValues = sortableData.map(item => item.row); // This is our final data
  Logger.log(`Sort complete. ${sortedValues.length} rows sorted.`);


  // --- 4. Create the NEW Sheet ---
  Logger.log("Attempting to delete old sorted sheet if it exists...");
  let sortedSheet = spreadsheet.getSheetByName(SORTED_SHEET_NAME);
  if (sortedSheet) {
    spreadsheet.deleteSheet(sortedSheet);
    SpreadsheetApp.flush(); // Make sure deletion is complete
    Logger.log("Old sheet deleted.");
  }
  
  Logger.log("Creating new sheet: " + SORTED_SHEET_NAME);
  sortedSheet = spreadsheet.insertSheet(SORTED_SHEET_NAME);
  spreadsheet.setActiveSheet(sortedSheet);
  Logger.log("New sheet created and activated.");


  // --- 5. Write Headers and Sorted Data to NEW Sheet ---
  if (sortedValues.length > 0) {
    Logger.log("Writing headers to new sheet...");
    // Write Headers
    sortedSheet.getRange(1, 1, 2, lastCol).setValues(headerValues);
    Logger.log("Writing sorted data to new sheet...");
    // Write Data
    sortedSheet.getRange(3, 1, sortedValues.length, lastCol).setValues(sortedValues);
    Logger.log("Data write complete. Flushing sheet.");
    SpreadsheetApp.flush(); // FORCE the sheet to update
  } else {
    Logger.log("Exiting: No data rows found to sort.");
    ui.alert("No data rows found to sort.");
    return;
  }

  // --- 6. Format the NEW Sheet ---
  Logger.log("Formatting new sheet (freezing panes, hiding column)...");
  try {
    sortedSheet.setFrozenRows(2);
    sortedSheet.setFrozenColumns(1);
    const sortKeyColNum = colIndices.sortKey + 1;
    sortedSheet.hideColumns(sortKeyColNum);
  } catch (e) {
     Logger.log("Error formatting new sheet: " + e.message);
  }
  
  // --- 7. Apply Highlights on NEW Sheet ---
  Logger.log("Applying highlights to new sheet...");
  let expansiveMarketRanges = [];
  let middleMarketRanges = [];
  let lowerPotentialRanges = [];
  const newDataRegion = sortedSheet.getRange(3, 1, sortedValues.length, lastCol);

  for (let i = 0; i < sortedValues.length; i++) {
    const rowNum = i + 3; // i=0 is Row 3
    const category = sortedValues[i][colIndices.sortKey]; 
    // Use sortedSheet here
    const rowRangeA1 = sortedSheet.getRange(rowNum, 1, 1, lastCol).getA1Notation(); 

    if (category === 1) expansiveMarketRanges.push(rowRangeA1);
    else if (category === 2) middleMarketRanges.push(rowRangeA1);
    else if (category === 3) lowerPotentialRanges.push(rowRangeA1);
  }

  // Clear all old backgrounds (on new sheet)
  newDataRegion.setBackground(null);
  
  // Apply formatting in batches (on new sheet)
  if (expansiveMarketRanges.length > 0) sortedSheet.getRangeList(expansiveMarketRanges).setBackground(EXPANSIVE_COLOR);
  if (middleMarketRanges.length > 0) sortedSheet.getRangeList(middleMarketRanges).setBackground(MIDDLE_MARKET_COLOR);
  if (lowerPotentialRanges.length > 0) sortedSheet.getRangeList(lowerPotentialRanges).setBackground(LOWER_POTENTIAL_COLOR);
  Logger.log("Highlighting complete.");

  // --- 8. Show final alert ---
  Logger.log("--- sortPotentials finished ---");
  ui.alert(`Sorting and highlighting complete! \n\nA new sheet named 'Sorted Market Analysis' has been created with ${expansiveMarketRanges.length} Expansive, ${middleMarketRanges.length} Medium Value, and ${lowerPotentialRanges.length} Lower Volume rows.`);
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